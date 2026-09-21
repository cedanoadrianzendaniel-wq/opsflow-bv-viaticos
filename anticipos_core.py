"""Generador PDF 'SOLICITUD DE ANTICIPO DE CONTABILIDAD' (Formato Autorizacion Descuento).

Rellena el PDF template BV con datos del trabajador + monto + fecha usando
overlay (reportlab canvas encima del PDF original).

Coordenadas calibradas con pdfplumber sobre el template BV oficial.
"""
import io, os, re
from datetime import date

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import white, black

TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
TEMPLATE_ANTICIPO = os.path.join(TEMPLATES_DIR, 'anticipo_bv.pdf')

_MESES = ['','ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']


def _reorder_nombre(nc):
    """NocoDB 'APELLIDO1 APELLIDO2 NOMBRE1 NOMBRE2' -> 'NOMBRE1 NOMBRE2 APELLIDO1 APELLIDO2'."""
    parts = (nc or '').split()
    if len(parts) >= 4:
        return ' '.join(parts[2:] + parts[:2])
    return nc or ''


def generate_anticipo_pdf(nombre_completo, dni, cargo, division, monto, fecha_solicitud=None):
    """Genera 1 PDF de solicitud de anticipo rellenado.

    Args:
      nombre_completo: str (viene ya en formato Nombres+Apellidos o Apellidos+Nombres,
        se reordena automaticamente si tiene 4+ tokens).
      dni: str
      cargo: str (puesto del trabajador)
      division: str (area/division, default 'Industria')
      monto: float o str (S/)
      fecha_solicitud: date | datetime | 'YYYY-MM-DD' | None (default hoy).
    Returns:
      bytes del PDF.
    """
    if fecha_solicitud is None:
        fecha_solicitud = date.today()
    elif isinstance(fecha_solicitud, str):
        y, m, d = fecha_solicitud[:10].split('-')
        fecha_solicitud = date(int(y), int(m), int(d))

    nombre_display = _reorder_nombre(nombre_completo)
    monto_str = f'{float(monto):,.2f}' if not isinstance(monto, str) else monto

    overlay_buf = io.BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=letter)  # US Letter 612x792
    c.setFont('Helvetica', 10)

    # Coordenadas calibradas: y_reportlab = 792 - top_pdfplumber + 7
    def y(top_pdfplumber):
        return 792 - top_pdfplumber + 7

    X = 205  # despues del label, sobre la linea de puntos
    c.drawString(X, y(198.3), nombre_display)
    c.drawString(X, y(232.9), str(dni or ''))
    c.drawString(X, y(267.3), str(cargo or ''))
    c.drawString(X, y(301.9), str(division or 'Industria'))
    c.drawString(X, y(388.2), monto_str)

    # Fecha: DD (x213) / MM (x262) / 202X - tapa el '202' del template y escribe año completo
    y_fecha = y(422.6)
    c.drawString(213, y_fecha, f'{fecha_solicitud.day:02d}')
    c.drawString(262, y_fecha, f'{fecha_solicitud.month:02d}')
    c.setFillColor(white)
    c.rect(287, y_fecha - 2, 22, 12, fill=1, stroke=0)
    c.setFillColor(black)
    c.drawString(289, y_fecha, str(fecha_solicitud.year))

    c.showPage()
    c.showPage()  # pagina 2 vacia (firma trabajador)
    c.save()
    overlay_buf.seek(0)

    template = PdfReader(TEMPLATE_ANTICIPO)
    overlay = PdfReader(overlay_buf)
    writer = PdfWriter()
    for i, page in enumerate(template.pages):
        if i < len(overlay.pages):
            page.merge_page(overlay.pages[i])
        writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def generate_anticipos_batch(workers_with_personal, fecha_solicitud=None):
    """Genera 1 PDF por trabajador matcheado (con monto > 0).

    workers_with_personal: [{'worker': {..., 'total'}, 'personal': {..., 'dni', 'nombre_completo', 'puesto', 'division'}}, ...]

    Returns:
      list de dicts {'filename', 'pdf_bytes', 'dni', 'nombre', 'monto'}.
    """
    out = []
    for item in workers_with_personal:
        p = item.get('personal') or {}
        w = item.get('worker') or {}
        if not p: continue
        monto = float(w.get('total', 0) or 0)
        if monto <= 0: continue
        dni = str(p.get('dni', ''))
        nombre = p.get('nombre_completo', '') or ''
        pdf = generate_anticipo_pdf(
            nombre_completo=nombre,
            dni=dni,
            cargo=p.get('puesto', '') or w.get('cargo', ''),
            division=p.get('division', '') or 'Industria',
            monto=monto,
            fecha_solicitud=fecha_solicitud,
        )
        safe = re.sub(r'[^A-Za-z0-9_-]', '_', nombre)[:40]
        out.append({
            'filename': f'DJ_Anticipo_{dni}_{safe}.pdf',
            'pdf_bytes': pdf,
            'dni': dni,
            'nombre': nombre,
            'monto': monto,
        })
    return out
