# Aquaponics dashboard backend — needs LibreOffice for live workbook recalculation.
# LibreOffice ships its Python UNO bridge (`uno`) only in the SYSTEM python, so we
# install all pip packages into that same system python to avoid two interpreters.
FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        python3-uno \
        python3-pip \
        fonts-dejavu \
        && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install FastAPI/uvicorn into the SYSTEM python3 (the one that can import uno).
COPY requirements.txt .
RUN python3 -m pip install --no-cache-dir --break-system-packages -r requirements.txt

COPY backend.py .
COPY Model_solver_multispecies_Parameters.xlsx .

# LibreOffice needs a writable HOME for its profile.
ENV HOME=/tmp
ENV MODEL_XLSX=/app/Model_solver_multispecies_Parameters.xlsx

# Render provides $PORT. Use the system python3 so `import uno` works.
CMD ["sh", "-c", "python3 -m uvicorn backend:app --host 0.0.0.0 --port ${PORT:-8000}"]
