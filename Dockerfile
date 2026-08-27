# syntax=docker/dockerfile:1
FROM node:22-bookworm AS ui
WORKDIR /src/app
COPY app/package.json app/package-lock.json* ./
RUN npm install
COPY app/ ./
RUN npm run build

FROM python:3.12-slim-bookworm
WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    INCI_PPWR_WEB=1 \
    INCI_PPWR_VERSION=1.0.0 \
    INCI_PPWR_WORKSPACE_ROOT=/data/workspace \
    INCI_PPWR_DELIVERY_ROOT=/data/delivery \
    INCI_PPWR_CANDIDATES_ROOT=/data/candidates \
    INCI_PPWR_ADMIN_USER=admin

RUN apt-get update && apt-get install -y --no-install-recommends \
      ca-certificates \
      libreoffice-writer \
      libreoffice-calc \
      fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY backend ./backend
COPY scripts ./scripts
COPY templates ./templates
COPY data_reference ./data_reference
COPY assets ./assets
COPY app/config ./app/config
COPY --from=ui /src/app/dist ./app/dist
COPY docker/entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Seed copies (Render disk overlays /data). Delivery is optional in git builds.
COPY workspace /opt/ppwr/workspace
COPY candidates /opt/ppwr/candidates
COPY delivery /opt/ppwr/delivery

RUN mkdir -p /data/workspace /data/delivery /data/candidates /opt/ppwr/delivery \
    && printf '%s\n' \
      '{' \
      '  "deliveryRoot": "/data/delivery",' \
      '  "candidatesRoot": "/data/candidates",' \
      '  "scopes": {' \
      '    "starter": "01_STARTER_INDIVIDUAL_DELIVERY_REV00",' \
      '    "industrial": "02_INDUSTRIAL_DELIVERY_REV00",' \
      '    "container": "03_CONTAINER_DELIVERY_REV00",' \
      '    "component": "04_COMPONENT_SPARE_DELIVERY_REV00"' \
      '  },' \
      '  "readOnlyDeliveries": true,' \
      '  "publishDate": "18.08.2026"' \
      '}' > ./app/config/app.config.json

WORKDIR /app/backend
EXPOSE 10000
ENTRYPOINT ["/entrypoint.sh"]
