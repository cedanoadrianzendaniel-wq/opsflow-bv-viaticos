"""Generador del F-004 Registro de Entrega de EPP.

Carga templates/epp_template.xlsx (con 17 EPPs hardcoded en filas 12-28) y rellena:
- Header (division, proyecto, nombre, dni, posicion)
- Tabla de EPPs: solo las filas correspondientes a los items entregados (resto queda en blanco)
- Observaciones
- Footer (Bureau Veritas + cliente)
"""
import io, os, re, unicodedata
from datetime import datetime
from openpyxl import load_workbook

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'templates', 'epp_template.xlsx')

# Lista oficial de 17 EPPs (orden de filas en el template: row 12 -> EPP_LIST[0], etc)
EPP_LIST = [
    'Casco',
    'Barbiquejo',
    'Camisa manga larga ignifugo',
    'Pantalos ignifugo',          # ojo: typo del template original 'Pantalos' (no Pantalon)
    'Casaca',
    'Lentes claros',
    'Lentes oscuros',
    'Protector auditivo tipo copa',
    'Chaleco anaranjado',
    'Polo M/L',
    'Botas cana alta',
    'Baston Treckin',
    'Capotin',
    'Botas de PVC',
    'Guantes de badana',
    'Camisa Oxford',
    'Pantalon Jean',
]
# Row 12 -> index 0, row 28 -> index 16
EPP_ROW_BASE = 12


def _norm(s: str) -> str:
    if not s: return ''
    s = unicodedata.normalize('NFD', str(s))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = re.sub(r'[^a-z0-9 ]', ' ', s.lower())
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def find_epp_row(nombre_epp: str):
    """Busca el row del template para un EPP. Devuelve row int o None."""
    target = _norm(nombre_epp)
    if not target: return None
    # Match exacto normalizado
    for idx, name in enumerate(EPP_LIST):
        if _norm(name) == target:
            return EPP_ROW_BASE + idx
    # Match por tokens: todos los tokens del target deben aparecer
    target_tokens = set(target.split())
    if not target_tokens: return None
    for idx, name in enumerate(EPP_LIST):
        name_tokens = set(_norm(name).split())
        # target_tokens subset de name_tokens, o name_tokens subset de target_tokens
        if target_tokens.issubset(name_tokens) or name_tokens.issubset(target_tokens):
            return EPP_ROW_BASE + idx
    # Match parcial: al menos 1 token coincide y es palabra "clave"
    keywords = {
        'casco': 0, 'barbiquejo': 1, 'camisa': None, 'pantalon': None, 'pantalos': 3,
        'casaca': 4, 'lentes': None, 'protector': 7, 'auditivo': 7, 'chaleco': 8,
        'polo': 9, 'botas': None, 'baston': 11, 'treckin': 11, 'trekking': 11,
        'capotin': 12, 'guantes': 14, 'badana': 14, 'oxford': 15, 'jean': 16,
    }
    for tok in target_tokens:
        idx = keywords.get(tok)
        if idx is not None:
            return EPP_ROW_BASE + idx
    # Camisa: ignifugo vs oxford
    if 'camisa' in target_tokens:
        if 'oxford' in target_tokens:
            return EPP_ROW_BASE + 15
        return EPP_ROW_BASE + 2  # ignifugo (default)
    # Pantalon: ignifugo vs jean
    if 'pantalon' in target_tokens or 'pantalones' in target_tokens:
        if 'jean' in target_tokens or 'jeans' in target_tokens:
            return EPP_ROW_BASE + 16
        return EPP_ROW_BASE + 3
    # Lentes: claros vs oscuros
    if 'lentes' in target_tokens:
        if 'oscuro' in target_tokens or 'oscuros' in target_tokens:
            return EPP_ROW_BASE + 6
        return EPP_ROW_BASE + 5
    # Botas: cana alta / pvc
    if 'botas' in target_tokens:
        if 'pvc' in target_tokens:
            return EPP_ROW_BASE + 13
        return EPP_ROW_BASE + 10
    return None


def _format_date(v):
    """Acepta datetime / 'YYYY-MM-DD' / 'DD/MM/YYYY' / None. Devuelve datetime o None."""
    if not v: return None
    if isinstance(v, datetime): return v
    s = str(v).strip()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def generate_epp_excel(trabajador: dict, items: list, cliente_data: dict = None, observaciones: str = None) -> bytes:
    """Genera el F-004 lleno y devuelve bytes del xlsx.

    trabajador: {nombre_completo, dni, puesto, cliente, division}
    items: [{nombre_epp, cantidad, fecha_entrega, fecha_cambio, talla, cantidad_pendiente, descripcion_cambio}, ...]
    cliente_data: opcional, info adicional del cliente (no usado en este formato basico)
    observaciones: texto libre para celda de observaciones
    """
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb['F-004']

    # Header
    ws['B5'] = trabajador.get('division') or 'INDUSTRIA'
    ws['B6'] = trabajador.get('cliente') or trabajador.get('proyecto') or ''
    ws['B7'] = trabajador.get('nombre_completo') or ''
    ws['B8'] = str(trabajador.get('dni') or '')
    ws['B9'] = trabajador.get('puesto') or trabajador.get('posicion') or ''

    # EPPs
    no_match = []
    for item in items:
        row = find_epp_row(item.get('nombre_epp') or item.get('nombre') or '')
        if row is None:
            no_match.append(item.get('nombre_epp') or item.get('nombre') or '?')
            continue
        if item.get('cantidad') is not None:
            ws[f'B{row}'] = item['cantidad']
        fe = _format_date(item.get('fecha_entrega'))
        if fe:
            ws[f'C{row}'] = fe
            ws[f'C{row}'].number_format = 'dd/mm/yyyy'
        fc = _format_date(item.get('fecha_cambio'))
        if fc:
            ws[f'D{row}'] = fc
            ws[f'D{row}'].number_format = 'dd/mm/yyyy'
        if item.get('talla'):
            ws[f'E{row}'] = item['talla']
        if item.get('cantidad_pendiente') is not None:
            ws[f'F{row}'] = item['cantidad_pendiente']
        if item.get('descripcion_cambio'):
            ws[f'G{row}'] = item['descripcion_cambio']

    # Observaciones: A29 es anchor de merge A29:I29 con label "OBSERVACIONES: ". Append valor.
    if observaciones:
        ws['A29'] = f'OBSERVACIONES: {observaciones}'

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), {'no_match_epps': no_match, 'matched_count': len(items) - len(no_match)}
