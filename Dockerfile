# APSG Report — Dockerfile
# NOTE: On Render Free Tier, use the native Python environment (render.yaml env: python)
# NOT Docker — Docker requires a paid plan for the Start Command to work correctly.
# This Dockerfile is provided for self-hosted / VPS / Docker Compose deployments.

FROM python:3.11-slim

WORKDIR /app

# Minimal system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies (cached layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Create required directories
RUN mkdir -p uploads outputs data

# Port (default 10000, overridden by $PORT env var)
EXPOSE 10000

# Free-tier optimised gunicorn startup
CMD ["sh", "-c", \
  "python -c 'from app import init_db; init_db()' && \
   gunicorn app:app \
     --bind 0.0.0.0:${PORT:-10000} \
     --workers 1 \
     --threads 1 \
     --timeout 900 \
     --max-requests 500 \
     --max-requests-jitter 50 \
     --worker-class sync \
     --log-level info"]
