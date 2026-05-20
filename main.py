"""FastAPI microservice para procesar viaticos.

Endpoint: POST /procesar_viaticos (multipart)
  - file: Excel detalle de viaticos
  - cliente: TGP / TDP / APM / PPC-PLUSPETROL / COGA
  - mes_label: YYYY-MM (ej "2026-05")

Devuelve un ZIP con:
  - Consolidado_Viaticos_{cliente}_{mes}.xlsx
  - Macro_SCT_Soles_{cliente}_{mes}.xlsm
  - result.json (metadata: matched, no_match, sin_cci, total_monto)
"""
import os, io, json, zipfile, urllib.request
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from viaticos_core import process_viaticos
from epps_core import generate_epp_excel, EPP_LIST

app = FastAPI(title="OpsFlow BV - Viaticos & EPPs Service", version="1.1")

# Config via env
NOCO_API_TOKEN = os.environ.get('NOCO_API_TOKEN', '')
NOCO_BASE = os.environ.get('NOCO_BASE', 'https://nocodb.opsflow.pe/api/v2')
TABLE_PERSONAL = os.environ.get('TABLE_PERSONAL', 'm0d8sxpntpga4ax')
TABLE_CLIENTES = os.environ.get('TABLE_CLIENTES', 'me9wb9xk2p4pxzx')

def http_get(url):
    req = urllib.request.Request(url, headers={'xc-token': NOCO_API_TOKEN})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode('utf-8'))

@app.get("/health")
def health():
    return {"status": "ok", "service": "viaticos-bv"}

@app.post("/procesar_viaticos")
async def procesar(
    file: UploadFile = File(...),
    cliente: str = Form(...),
    mes_label: str = Form(...),  # YYYY-MM
):
    if not NOCO_API_TOKEN:
        raise HTTPException(500, "NOCO_API_TOKEN no configurado")
    try:
        # Fetch personal y clientes desde NocoDB
        personal_resp = http_get(f'{NOCO_BASE}/tables/{TABLE_PERSONAL}/records?limit=200&fields=Id,nombre_completo,dni,puesto,banco,tipo_cuenta_abono,cuenta_cci,tipo_moneda,tipo_documento,cliente,division,correo_personal,correo_corporativo')
        personal_list = personal_resp.get('list', [])
        clientes_resp = http_get(f'{NOCO_BASE}/tables/{TABLE_CLIENTES}/records?limit=20')
        clientes_list = clientes_resp.get('list', [])
        cliente_data = next((c for c in clientes_list if c.get('cliente') == cliente), None)
        if not cliente_data:
            raise HTTPException(400, f"Cliente '{cliente}' no encontrado en clientes_bv. Disponibles: {[c.get('cliente') for c in clientes_list]}")

        # Leer archivo
        content = await file.read()

        # Procesar
        result = process_viaticos(content, cliente, mes_label, personal_list, cliente_data)

        # Empaquetar como ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr(f'Consolidado_Viaticos_{cliente}_{mes_label}.xlsx', result['consolidado_xlsx'])
            z.writestr(f'Macro_SCT_Soles_{cliente}_{mes_label}.xlsm', result['macro_xlsm'])
            z.writestr('result.json', json.dumps(result['metadata'], ensure_ascii=False, indent=2))
        zip_buf.seek(0)

        return StreamingResponse(
            zip_buf,
            media_type='application/zip',
            headers={
                'Content-Disposition': f'attachment; filename="viaticos_{cliente}_{mes_label}.zip"',
                'X-Total-Trabajadores': str(result['metadata']['matched']),
                'X-Total-Monto': str(result['metadata']['total_monto']),
                'X-No-Match-Count': str(len(result['metadata']['no_match'])),
            }
        )
    except HTTPException: raise
    except Exception as e:
        import traceback
        return JSONResponse(status_code=500, content={'error': str(e), 'trace': traceback.format_exc()[:2000]})


@app.post("/procesar_viaticos_json")
async def procesar_json(
    file: UploadFile = File(...),
    cliente: str = Form(...),
    mes_label: str = Form(...),
):
    """Variante: devuelve JSON con archivos en base64. Util si n8n no maneja bien zip."""
    import base64
    if not NOCO_API_TOKEN:
        raise HTTPException(500, "NOCO_API_TOKEN no configurado")
    personal_list = http_get(f'{NOCO_BASE}/tables/{TABLE_PERSONAL}/records?limit=200&fields=Id,nombre_completo,dni,puesto,banco,tipo_cuenta_abono,cuenta_cci,tipo_moneda,tipo_documento,cliente,division,correo_personal,correo_corporativo').get('list', [])
    clientes_list = http_get(f'{NOCO_BASE}/tables/{TABLE_CLIENTES}/records?limit=20').get('list', [])
    cliente_data = next((c for c in clientes_list if c.get('cliente') == cliente), None)
    if not cliente_data:
        raise HTTPException(400, f"Cliente '{cliente}' no encontrado")
    content = await file.read()
    result = process_viaticos(content, cliente, mes_label, personal_list, cliente_data)
    return {
        'consolidado_xlsx_b64': base64.b64encode(result['consolidado_xlsx']).decode('ascii'),
        'macro_xlsm_b64': base64.b64encode(result['macro_xlsm']).decode('ascii'),
        'consolidado_filename': f'Consolidado_Viaticos_{cliente}_{mes_label}.xlsx',
        'macro_filename': f'Macro_SCT_Soles_{cliente}_{mes_label}.xlsm',
        'metadata': result['metadata'],
    }


# ============================================================
# EPP - F-004 Registro de Entrega de EPP
# ============================================================
from fastapi import Body
from pydantic import BaseModel
from typing import List, Optional

class EppItem(BaseModel):
    nombre_epp: str
    cantidad: Optional[int] = None
    fecha_entrega: Optional[str] = None  # YYYY-MM-DD o DD/MM/YYYY
    fecha_cambio: Optional[str] = None
    talla: Optional[str] = None
    cantidad_pendiente: Optional[int] = None
    descripcion_cambio: Optional[str] = None

class GenerarEppRequest(BaseModel):
    dni: Optional[str] = None
    nombre_trabajador: Optional[str] = None  # si no hay dni, intenta match por nombre
    observaciones: Optional[str] = None
    items: List[EppItem]


@app.get("/epp/lista_oficial")
def epp_lista():
    """Devuelve la lista oficial de los 17 EPPs (orden = row del template)."""
    return {'epps': EPP_LIST}


@app.post("/generar_epp")
def generar_epp(payload: GenerarEppRequest):
    """Genera el F-004 Registro de Entrega de EPP para un trabajador.

    Busca al trabajador en personal_bv por DNI (preferido) o nombre.
    Devuelve el xlsx en base64 + metadata.
    """
    import base64
    if not NOCO_API_TOKEN:
        raise HTTPException(500, "NOCO_API_TOKEN no configurado")

    # Buscar trabajador
    trab = None
    if payload.dni:
        resp = http_get(f'{NOCO_BASE}/tables/{TABLE_PERSONAL}/records?where=(dni,eq,{payload.dni})&limit=1')
        if resp.get('list'):
            trab = resp['list'][0]
    if not trab and payload.nombre_trabajador:
        # Fuzzy: tokens del nombre todos en personal.nombre_completo
        nombre_norm = payload.nombre_trabajador.upper()
        tokens = [t for t in nombre_norm.split() if len(t) >= 3]
        if tokens:
            where = '~and'.join([f'(nombre_completo,like,%25{t}%25)' for t in tokens])
            resp = http_get(f'{NOCO_BASE}/tables/{TABLE_PERSONAL}/records?where={where}&limit=5')
            if len(resp.get('list', [])) == 1:
                trab = resp['list'][0]
            elif len(resp.get('list', [])) > 1:
                return JSONResponse(status_code=400, content={
                    'error': 'multiples_matches',
                    'matches': [{'Id': p['Id'], 'nombre': p['nombre_completo'], 'dni': p.get('dni')} for p in resp['list']]
                })
    if not trab:
        raise HTTPException(404, f"Trabajador no encontrado (dni={payload.dni}, nombre={payload.nombre_trabajador})")

    items = [it.model_dump() for it in payload.items]
    xlsx_bytes, meta = generate_epp_excel(trab, items, observaciones=payload.observaciones)

    return {
        'xlsx_b64': base64.b64encode(xlsx_bytes).decode('ascii'),
        'filename': f'EPP_{(trab.get("cliente") or "BV")}-{trab.get("nombre_completo","").replace(" ","_")}.xlsx',
        'trabajador': {
            'Id': trab.get('Id'),
            'dni': trab.get('dni'),
            'nombre_completo': trab.get('nombre_completo'),
            'cliente': trab.get('cliente'),
            'division': trab.get('division'),
            'puesto': trab.get('puesto'),
            'correo_personal': trab.get('correo_personal'),
            'correo_corporativo': trab.get('correo_corporativo'),
            'telefono': trab.get('telefono'),
        },
        'meta': meta,
    }
