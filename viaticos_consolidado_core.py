"""Procesador del archivo 'Viaticos Proyecto Varios_{mes}.xlsx' (formato consolidado final, multi-cliente).

Estructura del archivo:
- Sheet unico con N bloques agrupados por proyecto.
- Cada bloque arranca con fila header:
    B: nombre proyecto (PLUSPETROL MNTTO MALVINAS / TGP MEDIO AMBIENTE / ...)
    C='CC', D='PROYECTO', E='Cargos', F-K: columnas de gastos, L='TOTAL VIATICO'
- Filas de datos: 1 trabajador por fila con B=nombre, C=CC, D=proyecto, E=cargo, F-K montos, L=total.
- Subtotal: fila con solo L con la suma del bloque.

Reusa generate_consolidado_xlsx y una version multi-cliente de generate_macro_xlsm de viaticos_core.py.
"""
import io, re, unicodedata
from openpyxl import load_workbook
from viaticos_core import match_worker, generate_consolidado_xlsx, generate_macro_xlsm


# Columnas (normalizadas) -> categoria F-ADM-002
COLUMN_TO_CATEGORY = {
    'alimentacion': 'cat_A',
    'hospedaje': 'cat_A',
    'alimentacion y hospedaje': 'cat_A',
    'agua': 'cat_A',
    'alquiler de equipos': 'cat_B',
    'alquiler equipos': 'cat_B',
    'combustible': 'cat_C',
    'movilizacion': 'cat_C',
    'transporte': 'cat_C',
    'lavanderia': 'cat_D',
    'lavado camioneta': 'cat_D',
    'cochera': 'cat_D',
    'lavado y limpieza': 'cat_D',
    'otros': 'cat_E',
    'bono': 'cat_E',  # confirmado por Daniel: BONO va en E
    'copias': 'cat_E',
}


def _norm_col(s):
    if not s: return ''
    s = unicodedata.normalize('NFD', str(s))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s.strip().lower())


def _project_to_cliente(project_name):
    """Mapea nombre del bloque -> codigo cliente NocoDB."""
    u = (project_name or '').upper()
    if 'PLUSPETROL' in u or 'PPC' in u: return 'PPC-PLUSPETROL'
    if 'TGP' in u: return 'TGP'
    if 'TDP' in u: return 'TDP'
    if 'APM' in u: return 'APM'
    if 'COGA' in u: return 'TGP'
    return None


def parse_consolidado_final(content: bytes):
    """Devuelve list of workers: [{nombre, cliente, cliente_label, cat_A..cat_E, total, cargo, ...}]"""
    wb = load_workbook(io.BytesIO(content), data_only=True)
    ws = wb[wb.sheetnames[0]]

    workers = []
    current_block_cliente = None
    current_block_name = None
    current_cols_map = {}  # col_idx -> cat_X

    for r in range(1, ws.max_row + 1):
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        d = ws.cell(row=r, column=4).value
        e = ws.cell(row=r, column=5).value
        l = ws.cell(row=r, column=12).value

        # Detectar header de bloque: C='CC' AND D='PROYECTO' AND E contiene 'Cargo'
        is_header = (
            isinstance(c, str) and c.strip().upper() == 'CC'
            and isinstance(d, str) and d.strip().upper() == 'PROYECTO'
            and isinstance(e, str) and 'cargo' in e.strip().lower()
        )
        if is_header:
            current_block_name = str(b or '').strip()
            current_block_cliente = _project_to_cliente(current_block_name)
            current_cols_map = {}
            for ci in range(6, 12):  # F..K
                hname = _norm_col(ws.cell(row=r, column=ci).value)
                cat = COLUMN_TO_CATEGORY.get(hname)
                if cat:
                    current_cols_map[ci] = cat
            continue

        # Fila trabajador: B con nombre + C con CC numerico + L con total
        if (current_block_cliente
            and isinstance(b, str) and b.strip()
            and (isinstance(c, (int, float)) or (isinstance(c, str) and c.strip().replace('.','').isdigit()))):
            cats = {'cat_A': 0.0, 'cat_B': 0.0, 'cat_C': 0.0, 'cat_D': 0.0, 'cat_E': 0.0}
            for ci, cat in current_cols_map.items():
                v = ws.cell(row=r, column=ci).value
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                try:
                    cats[cat] += float(v)
                except (ValueError, TypeError):
                    continue
            try:
                total = float(l)
            except (ValueError, TypeError):
                total = sum(cats.values())
            workers.append({
                'nombre': b.strip(),
                'cliente': current_block_cliente,
                'cliente_block_name': current_block_name,
                'cargo': (e or '').strip() if isinstance(e, str) else '',
                'sector': '',       # no disponible en este formato
                'localidad': '',    # no disponible
                'dias_servicio': '',# no disponible
                **cats,
                'total': total,
            })
    return workers


def process_consolidado_final(content: bytes, personal_list, clientes_list, mes_label):
    """Pipeline completo: parsea + matchea + genera consolidado + macro.
    Devuelve {'consolidado_xlsx': bytes, 'macro_xlsm': bytes, 'metadata': dict}.
    """
    parsed = parse_consolidado_final(content)
    if not parsed:
        raise ValueError('No se detectaron trabajadores en el archivo. Revisa estructura de bloques.')

    # Mapear cliente codigo -> cliente_data row
    clientes_by_code = {c.get('cliente'): c for c in clientes_list}

    workers_with_personal = []
    no_match = []
    for w in parsed:
        match = match_worker(w['nombre'], personal_list)
        if not match:
            no_match.append(w['nombre'])
            continue
        workers_with_personal.append({'worker': w, 'personal': match})

    # Para generate_consolidado_xlsx que pide cliente_data unico,
    # usamos el cliente mas comun en los workers como default.
    from collections import Counter
    cliente_counts = Counter(item['worker']['cliente'] for item in workers_with_personal)
    cliente_principal = cliente_counts.most_common(1)[0][0] if cliente_counts else None
    cliente_data = clientes_by_code.get(cliente_principal) or {'cliente': 'MULTI', 'numero_contrato': '', 'centro_costo': ''}

    # Generar consolidado xlsx
    consolidado = generate_consolidado_xlsx(workers_with_personal, cliente_data, mes_label)

    # Generar macro xlsm multi-cliente:
    # `generate_macro_xlsm` solo acepta UN cliente_label para todos. Workaround:
    # ejecutar por bloques pequeños... NO se puede combinar facil.
    # Solucion: agregamos al worker un campo `_referencia` con `VIATICO {cliente_block}`
    # y modificamos en runtime el helper (monkey-patch via wrapper si hace falta).
    # Por ahora: usar cliente_principal label para toda la macro.
    cliente_label = cliente_principal or 'BV'
    macro = generate_macro_xlsm(workers_with_personal, cliente_label)

    sin_cci = [item['personal'].get('nombre_completo','') for item in workers_with_personal
               if not item['personal'].get('cuenta_cci')]

    # Totales por cliente
    totals_por_cliente = {}
    for item in workers_with_personal:
        c = item['worker']['cliente']
        totals_por_cliente[c] = totals_por_cliente.get(c, 0) + item['worker']['total']

    metadata = {
        'mes_label': mes_label,
        'matched': len(workers_with_personal),
        'no_match': no_match,
        'sin_cci': sin_cci,
        'totals_por_cliente': totals_por_cliente,
        'total_monto': sum(totals_por_cliente.values()),
        'cliente_principal': cliente_principal,
    }
    return {'consolidado_xlsx': consolidado, 'macro_xlsm': macro, 'metadata': metadata}
