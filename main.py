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
from epps_core import generate_epp_excel, EPP_LIST, parse_excel_masivo
from viaticos_consolidado_core import process_consolidado_final, parse_consolidado_final
import boto3
from botocore.client import Config as BotoConfig

# MinIO config
MINIO_ENDPOINT = os.environ.get('MINIO_ENDPOINT', 'https://control-vacaciones-bv-minio.dxgsgp.easypanel.host')
MINIO_ACCESS_KEY = os.environ.get('MINIO_ACCESS_KEY', 'bvopsflow')
MINIO_SECRET_KEY = os.environ.get('MINIO_SECRET_KEY', 'BureauVeritas2026ASIS')
MINIO_BUCKET = os.environ.get('MINIO_BUCKET', 'prorrogas')

def _s3_client():
    return boto3.client(
        's3',
        endpoint_url=MINIO_ENDPOINT,
        aws_access_key_id=MINIO_ACCESS_KEY,
        aws_secret_access_key=MINIO_SECRET_KEY,
        config=BotoConfig(signature_version='s3v4', s3={'addressing_style': 'path'}),
        region_name='us-east-1',
    )

def _claude_match_worker(nombre_input, personal_list):
    """Fallback con criterio: cuando el match por tokens falla, Claude resuelve
    el nombre contra la lista de personal (maneja orden distinto, comas, nombres
    parciales, typos). Devuelve el dict del trabajador o None."""
    import re as _re
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return None
    try:
        from anthropic import Anthropic
        client = Anthropic(api_key=api_key)
        candidatos = [{'id': p.get('Id'), 'nombre': p.get('nombre_completo'),
                       'dni': p.get('dni'), 'cliente': p.get('cliente')} for p in personal_list]
        prompt = (
            'Tengo un nombre de trabajador escrito en un Excel y debo encontrar a que persona '
            'de una lista de personal corresponde. Considera que el ORDEN de nombres y apellidos '
            'puede variar, puede haber comas, abreviaciones, segundos nombres faltantes, o pequenos typos. '
            'Match por la persona, no por coincidencia exacta de texto.\n\n'
            f'Nombre del Excel: "{nombre_input}"\n\n'
            f'Lista de personal (JSON): {json.dumps(candidatos, ensure_ascii=False)}\n\n'
            'Responde SOLO JSON sin markdown: {"id": <Id del match o null>, "confianza": "alta|media|baja"}. '
            'Si no hay una persona que claramente sea la misma, id=null.'
        )
        msg = client.messages.create(
            model='claude-haiku-4-5-20251001', max_tokens=100,
            messages=[{'role': 'user', 'content': prompt}]
        )
        raw = msg.content[0].text.strip()
        raw = _re.sub(r'^```(?:json)?\s*', '', raw, flags=_re.IGNORECASE)
        raw = _re.sub(r'```\s*$', '', raw).strip()
        dec = json.loads(raw)
        if dec.get('id') and dec.get('confianza') in ('alta', 'media'):
            for p in personal_list:
                if p.get('Id') == dec['id']:
                    print(f'[CLAUDE MATCH] "{nombre_input}" -> {p.get("nombre_completo")} (conf {dec.get("confianza")})')
                    return p
        print(f'[CLAUDE MATCH] "{nombre_input}" -> sin match (id={dec.get("id")} conf={dec.get("confianza")})')
    except Exception as e:
        print('[CLAUDE MATCH] error:', e)
    return None


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


@app.get("/epp/debug")
def epp_debug():
    """Diagnostico: existe LOGO_PATH? que archivos hay en templates/?"""
    import os as _os
    from epps_core import LOGO_PATH, TEMPLATE_PATH
    tpath_dir = _os.path.dirname(TEMPLATE_PATH)
    files = []
    if _os.path.isdir(tpath_dir):
        for f in sorted(_os.listdir(tpath_dir)):
            full = _os.path.join(tpath_dir, f)
            files.append({'name': f, 'size': _os.path.getsize(full), 'is_file': _os.path.isfile(full)})
    return {
        'logo_path': LOGO_PATH,
        'logo_exists': _os.path.exists(LOGO_PATH),
        'logo_size': _os.path.getsize(LOGO_PATH) if _os.path.exists(LOGO_PATH) else None,
        'template_dir': tpath_dir,
        'template_files': files,
    }


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


@app.post("/procesar_viaticos_consolidado_final")
async def procesar_viaticos_consolidado_final_endpoint(
    file: UploadFile = File(...),
    mes_label: str = Form(...),  # YYYY-MM
):
    """Procesa el archivo 'Viaticos Proyecto Varios_{mes}.xlsx' (multi-cliente).
    Devuelve JSON con consolidado + macro en base64.
    """
    import base64
    if not NOCO_API_TOKEN:
        raise HTTPException(500, "NOCO_API_TOKEN no configurado")
    try:
        personal_list = http_get(f'{NOCO_BASE}/tables/{TABLE_PERSONAL}/records?limit=500&fields=Id,nombre_completo,dni,puesto,banco,tipo_cuenta_abono,cuenta_cci,tipo_moneda,tipo_documento,cliente,division,correo_personal,correo_corporativo').get('list', [])
        clientes_list = http_get(f'{NOCO_BASE}/tables/{TABLE_CLIENTES}/records?limit=20').get('list', [])

        content = await file.read()
        result = process_consolidado_final(content, personal_list, clientes_list, mes_label)

        return {
            'consolidado_xlsx_b64': base64.b64encode(result['consolidado_xlsx']).decode('ascii'),
            'macro_xlsm_b64': base64.b64encode(result['macro_xlsm']).decode('ascii'),
            'consolidado_filename': f'Consolidado_Viaticos_MultiCliente_{mes_label}.xlsx',
            'macro_filename': f'Macro_SCT_Soles_MultiCliente_{mes_label}.xlsm',
            'metadata': result['metadata'],
        }
    except HTTPException: raise
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        import traceback
        return JSONResponse(status_code=500, content={'error': str(e), 'trace': traceback.format_exc()[:2000]})


@app.post("/procesar_epp_masivo")
async def procesar_epp_masivo(file: UploadFile = File(...)):
    """Procesa el Excel masivo 'Entrega de Epps {cliente}.xlsx' (1 fila = 1 trabajador).
    Para cada trabajador busca match en personal_bv (por nombre fuzzy) y genera el F-004.
    Devuelve JSON con N entries (cada una con xlsx_b64 + datos del trabajador + items + meta).
    """
    import base64
    if not NOCO_API_TOKEN:
        raise HTTPException(500, "NOCO_API_TOKEN no configurado")

    content = await file.read()
    try:
        filas = parse_excel_masivo(content)
    except ValueError as e:
        raise HTTPException(400, str(e))

    if not filas:
        raise HTTPException(400, "Excel vacio (no se detectaron trabajadores)")

    # Fetch personal completo (1 query)
    personal_resp = http_get(f'{NOCO_BASE}/tables/{TABLE_PERSONAL}/records?limit=500&fields=Id,nombre_completo,dni,puesto,cliente,division,correo_personal,correo_corporativo,telefono')
    personal_list = personal_resp.get('list', [])

    def find_trab(nombre_input: str):
        """Match fuzzy: tokens del input deben aparecer en personal.nombre_completo."""
        import unicodedata
        import re as _re
        def norm(s):
            s = unicodedata.normalize('NFD', str(s or '').upper())
            s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
            # Quitar puntuacion (comas, puntos, etc.) -> solo letras/numeros/espacios
            s = _re.sub(r'[^A-Z0-9 ]', ' ', s)
            s = _re.sub(r'\s+', ' ', s).strip()
            return s
        tokens_input = [t for t in norm(nombre_input).split() if len(t) >= 3]
        if not tokens_input: return None
        candidatos = []
        for p in personal_list:
            n_norm = norm(p.get('nombre_completo'))
            if all(t in n_norm for t in tokens_input):
                candidatos.append(p)
        if len(candidatos) == 1:
            return candidatos[0]
        if candidatos:
            # Si hay varios, devolver el primero (advertencia se incluye en meta)
            return candidatos[0]
        # Sin match por tokens -> fallback con criterio (Claude)
        return _claude_match_worker(nombre_input, personal_list)

    results = []
    no_match = []
    matched = 0
    for fila in filas:
        trab = find_trab(fila['nombre'])
        if not trab:
            no_match.append(fila['nombre'])
            results.append({
                'nombre_input': fila['nombre'],
                'matched': False,
                'reason': 'no_match_en_personal_bv',
            })
            continue
        # Override con datos del Excel si vienen
        trab_data = {**trab}
        if fila['puesto']:
            trab_data['puesto'] = fila['puesto']
        if fila['proyecto']:
            # Normalize proyecto a los valores canonicos validos en NocoDB
            # NocoDB single-select valida estrictamente: TGP, TDP, APM, PPC-PLUSPETROL, BACK OFFICE
            _p = str(fila['proyecto']).strip().upper()
            _CLI_MAP = {
                'PPC': 'PPC-PLUSPETROL', 'PLUSPETROL': 'PPC-PLUSPETROL', 'PPC-PLUSPETROL': 'PPC-PLUSPETROL',
                'TGP': 'TGP', 'COGA': 'TGP', 'TGP-COGA': 'TGP',
                'TDP': 'TDP', 'TERMINALES DEL PERU': 'TDP', 'TERMINALES': 'TDP', 'TDP-TERMINALES': 'TDP',
                'APM': 'APM', 'APM TERMINALS': 'APM', 'APM-TERMINALS': 'APM',
                'BACK OFFICE': 'BACK OFFICE', 'BACKOFFICE': 'BACK OFFICE', 'BO': 'BACK OFFICE',
            }
            trab_data['cliente'] = _CLI_MAP.get(_p, trab.get('cliente') or _p)

        xlsx_bytes, meta = generate_epp_excel(trab_data, fila['items'])
        # Filename limpio: sin comas, sin acentos, sin caracteres problematicos para URLs/Meta
        import re as _re, unicodedata as _ud
        _nm = trab_data.get("nombre_completo","")
        _nm = _ud.normalize('NFD', _nm)
        _nm = ''.join(c for c in _nm if _ud.category(c) != 'Mn')
        _nm = _re.sub(r'[^A-Za-z0-9 ]', '', _nm).strip()
        _nm = _re.sub(r'\s+', '_', _nm)
        filename_base = f'EPP_{(trab_data.get("cliente") or "BV")}-{_nm}'
        filename = f'{filename_base}.xlsx'
        filename_pdf = f'{filename_base}.pdf'

        # Convertir xlsx -> pdf con libreoffice (WhatsApp template no acepta xlsx)
        pdf_bytes = None
        try:
            import tempfile, subprocess, os as _os
            with tempfile.TemporaryDirectory() as tmpdir:
                xlsx_path = _os.path.join(tmpdir, filename)
                with open(xlsx_path, 'wb') as f:
                    f.write(xlsx_bytes)
                subprocess.run(
                    ['libreoffice', '--headless', '--convert-to', 'pdf',
                     '--outdir', tmpdir, xlsx_path],
                    check=True, capture_output=True, timeout=60
                )
                pdf_path = _os.path.join(tmpdir, filename_base + '.pdf')
                if _os.path.exists(pdf_path):
                    with open(pdf_path, 'rb') as f:
                        pdf_bytes = f.read()
        except Exception as _pdf_ex:
            meta = {**(meta or {}), 'pdf_conversion_error': str(_pdf_ex)[:200]}

        # Upload to MinIO + generate presigned URL (xlsx + pdf si esta disponible)
        import datetime as _dt
        lima = _dt.datetime.utcnow() - _dt.timedelta(hours=5)
        date_str = lima.strftime("%Y-%m-%d")
        ts = int(_dt.datetime.utcnow().timestamp()*1000)
        key = f'epps/{date_str}/{ts}_{matched}_{filename}'
        key_pdf = f'epps/{date_str}/{ts}_{matched}_{filename_pdf}'
        presigned_url_pdf = None
        url_public_pdf = None
        try:
            s3 = _s3_client()
            s3.put_object(
                Bucket=MINIO_BUCKET,
                Key=key,
                Body=xlsx_bytes,
                ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            presigned_url = s3.generate_presigned_url(
                'get_object',
                Params={'Bucket': MINIO_BUCKET, 'Key': key},
                ExpiresIn=604800,  # 7 días
            )
            url_public = f'{MINIO_ENDPOINT}/{MINIO_BUCKET}/{key}'
            # Upload PDF si esta disponible
            if pdf_bytes:
                s3.put_object(
                    Bucket=MINIO_BUCKET,
                    Key=key_pdf,
                    Body=pdf_bytes,
                    ContentType='application/pdf',
                )
                presigned_url_pdf = s3.generate_presigned_url(
                    'get_object',
                    Params={'Bucket': MINIO_BUCKET, 'Key': key_pdf},
                    ExpiresIn=604800,
                )
                url_public_pdf = f'{MINIO_ENDPOINT}/{MINIO_BUCKET}/{key_pdf}'
        except Exception as ex:
            results.append({
                'matched': False,
                'nombre_input': fila['nombre'],
                'reason': f'upload_minio_error: {ex}',
            })
            continue

        results.append({
            'matched': True,
            'nombre_input': fila['nombre'],
            'filename': filename,
            'filename_pdf': filename_pdf,
            'items_count': len(fila['items']),
            'url_minio_original': url_public,
            'url_minio_pdf': url_public_pdf,
            'presigned_url': presigned_url,
            'presigned_url_pdf': presigned_url_pdf,
            'trabajador': {
                'Id': trab.get('Id'),
                'dni': trab.get('dni'),
                'nombre_completo': trab.get('nombre_completo'),
                'cliente': trab_data.get('cliente'),
                'division': trab.get('division'),
                'puesto': trab_data.get('puesto'),
                'correo_personal': trab.get('correo_personal'),
                'correo_corporativo': trab.get('correo_corporativo'),
                'telefono': trab.get('telefono'),
            },
            'items': fila['items'],
            'meta': meta,
        })
        matched += 1

    return {
        'total_filas': len(filas),
        'matched': matched,
        'no_match': no_match,
        'results': results,
    }


@app.get("/onboarding/presigned_urls")
def onboarding_presigned_urls():
    """Devuelve URLs presignadas (7 dias) para los PDFs de onboarding en MinIO."""
    s3 = _s3_client()
    files = {
        'welcome': 'onboarding/Welcome_Colaborador_BV.pdf',
        'listado': 'onboarding/F-RRHH-091_Listado_Documentos_Ingresos.pdf',
    }
    urls = {}
    for k, key in files.items():
        urls[k] = s3.generate_presigned_url(
            'get_object',
            Params={'Bucket': MINIO_BUCKET, 'Key': key},
            ExpiresIn=604800,  # 7 dias
        )
    return urls
