# Aquaponics dashboard backend.
#
# LibreOffice's Python UNO bridge ("uno") is compiled against the SYSTEM python
# that ships with the distro's LibreOffice packages. Mixing it with a separately
# installed python (as python:3.x images do) causes: ModuleNotFoundError: No module
# named 'uno'. The reliable fix is to use ONE python — the distro's — for both
# LibreOffice and the web app. So we base on Ubuntu and use /usr/bin/python3
# throughout.
FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        python3-uno \
        python3 \
        python3-pip \
        procps \
        fonts-dejavu \
        && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install web deps into the SAME system python that can import uno.
COPY requirements.txt .
RUN python3 -m pip install --no-cache-dir --break-system-packages -r requirements.txt

COPY backend.py .
COPY Model_solver_multispecies_Parameters.xlsx .

# LibreOffice needs a writable HOME for its user profile.
ENV HOME=/tmp
ENV MODEL_XLSX=/app/Model_solver_multispecies_Parameters.xlsx

# Sanity check at build time: fail the build early if uno isn't importable.
RUN python3 -c "import uno; print('uno import OK')"

# Render provides $PORT.
CMD ["sh", "-c", "python3 -m uvicorn backend:app --host 0.0.0.0 --port ${PORT:-8000}"]
