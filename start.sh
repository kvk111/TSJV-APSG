#!/bin/bash
# APSG (Staging Ground) Report — Render Free Tier Startup Script
#
# Memory optimisation for 512 MB RAM / 0.1 CPU:
#   --workers 1       : single worker = one Python process = one copy of RAM
#   --threads 1       : single thread = sequential requests, no concurrent memory spikes
#   --timeout 900     : 15 min — allows heavy Excel/PPT builds to complete on slow CPU
#   --max-requests 500: restart worker after 500 requests to prevent memory leak buildup
#                       (was 50 — too aggressive; caused worker restart between process
#                        and build steps, wiping in-memory state and causing "Run /ct/process
#                        first" errors even though processing had succeeded)
#   --max-requests-jitter 50: stagger restarts gracefully
#   NO --preload      : avoids duplicating startup memory into worker at fork time

python -c "from app import init_db; init_db()"
exec gunicorn app:app \
    --bind "0.0.0.0:${PORT:-10000}" \
    --workers 1 \
    --threads 1 \
    --timeout 900 \
    --max-requests 500 \
    --max-requests-jitter 50 \
    --worker-class sync \
    --log-level info
