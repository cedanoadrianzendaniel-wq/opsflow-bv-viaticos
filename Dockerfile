FROM python:3.11-slim

WORKDIR /app

# LibreOffice headless para convertir XLSX -> PDF (necesario para WhatsApp template que rechaza xlsx)
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        libreoffice-core \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY templates/ /app/templates/
COPY viaticos_core.py epps_core.py viaticos_consolidado_core.py main.py /app/

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:3000/health', timeout=3)" || exit 1

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "3000"]
