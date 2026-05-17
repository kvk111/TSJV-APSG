"""
ct_module/__init__.py
Cycle Time Report — Flask Blueprint
Integrated into ABSC Main as module 05.
All routes are prefixed with /ct/ so they never collide with existing ABSC routes.

MEMORY OPTIMISATION (Render Free Tier — 512 MB / 0.1 CPU):
  - DataFrames are NEVER held in memory after report bytes are written.
  - Each build step writes its output to a named temp file, then clears the DF.
  - gc.collect() is called explicitly after each heavy step.
  - Reports are streamed directly from the temp file for download.
  - Only one report stage runs at a time (sequential, not parallel).
  - Presentations (PPT_DB) are cleared from memory after every PPT build.
  - The in-memory _ct_store holds ONLY lightweight metadata (scalars + file paths).
"""

import gc
import io
import os
import sys
import tempfile
import traceback
from datetime import datetime

from flask import (Blueprint, make_response, jsonify, redirect,
                   render_template_string, request, session, url_for)

# Make ct_module itself importable as a package path for its sibling modules
_CT_DIR = os.path.dirname(os.path.abspath(__file__))
if _CT_DIR not in sys.path:
    sys.path.insert(0, _CT_DIR)

ct = Blueprint('ct', __name__, url_prefix='/ct')

# ── Temp-file directory (Render ephemeral disk, wiped on restart) ──────────────
_TMP_DIR = tempfile.mkdtemp(prefix='apsg_ct_')

def _tmp_path(name):
    """Return a stable path inside our private temp dir."""
    return os.path.join(_TMP_DIR, name)

def _write_tmp(name, data: bytes) -> str:
    """Write bytes to a temp file and return its path."""
    path = _tmp_path(name)
    with open(path, 'wb') as f:
        f.write(data)
    return path

def _read_tmp(name) -> bytes | None:
    """Read bytes from a temp file, return None if it doesn't exist."""
    path = _tmp_path(name)
    if not os.path.exists(path):
        return None
    with open(path, 'rb') as f:
        return f.read()

def _del_tmp(*names):
    """Delete temp files; silently ignore missing ones."""
    for name in names:
        try:
            os.remove(_tmp_path(name))
        except FileNotFoundError:
            pass

def _gc():
    """Explicit garbage collection — reclaims pandas/openpyxl memory."""
    gc.collect()
    gc.collect()   # two passes catch cycle-referenced objects

# ── In-memory store: ONLY lightweight metadata ─────────────────────────────────
# Heavy data (DataFrames, xlsx bytes, ppt bytes) lives on disk in _TMP_DIR.
# This dict holds ONLY counts, scalars, dates, and flags.
_ct_store = {
    # Lightweight metadata
    'ct_records':      0,
    'wb_records':      0,
    'exceedances':     0,
    'applicable_hours': 0,
    'net_weight':      0.0,
    'failure_list':    [],    # list of 'fail'/'ok' — small for typical datasets
    'start_dt':        None,
    'end_dt':          None,
    'reason':          '',
    # Processing flags (set True when temp file exists)
    'processed':  False,
    'main_done':  False,
    'sum_done':   False,
    'ppt_done':   False,
}

# Temp-file name constants
_TF_CT_PARQUET   = 'ct_clean.parquet'
_TF_AN_PARQUET   = 'ct_anomaly.parquet'
_TF_WB_PARQUET   = 'wb_total.parquet'
_TF_RJ_PARQUET   = 'wb_rejected.parquet'
_TF_MAIN_XLSX    = 'main_report.xlsx'
_TF_SUMMARY_XLSX = 'summary_report.xlsx'
_TF_PPT          = 'ct_report.pptx'
_TF_META_JSON    = 'ct_meta.json'   # persists _ct_store scalars across worker restarts


def _save_meta():
    """
    Persist lightweight _ct_store metadata to disk as JSON.
    Called after every successful /ct/process so the state survives gunicorn
    worker restarts (--max-requests recycles the worker process, wiping globals).
    datetime objects are serialised as ISO strings; None becomes null.
    """
    import json as _json
    meta = {
        k: (v.isoformat() if hasattr(v, 'isoformat') else v)
        for k, v in _ct_store.items()
        if k not in ('failure_list',)   # failure_list saved separately if large
    }
    meta['failure_list'] = _ct_store.get('failure_list', [])
    path = _tmp_path(_TF_META_JSON)
    with open(path, 'w') as f:
        _json.dump(meta, f)


def _load_meta():
    """
    Reload _ct_store from disk if the in-memory store has been wiped by a
    worker restart (i.e. 'processed' is False but the JSON file exists on disk).
    Returns True if state was successfully restored, False if no saved state found.
    """
    import json as _json
    from datetime import datetime as _dt
    if _ct_store.get('processed'):
        return True                   # already live in memory — nothing to do
    path = _tmp_path(_TF_META_JSON)
    if not os.path.exists(path):
        return False
    try:
        with open(path) as f:
            meta = _json.load(f)
        # Re-parse datetime strings back to datetime objects
        for key in ('start_dt', 'end_dt'):
            if meta.get(key):
                try:
                    meta[key] = _dt.fromisoformat(meta[key])
                except (ValueError, TypeError):
                    meta[key] = None
        _ct_store.update(meta)
        return bool(_ct_store.get('processed'))
    except Exception:
        return False




def _get_ct_engine():
    """Lazy import of Cycle Time report_engine (ct_module version)."""
    import importlib.util
    mod_name = 'ct_report_engine'
    if mod_name in sys.modules:
        return sys.modules[mod_name]
    spec = importlib.util.spec_from_file_location(
        mod_name, os.path.join(_CT_DIR, 'report_engine.py'))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod


def _get_ct_app():
    """Lazy import of Cycle Time app module (for build_full_report etc.)."""
    import importlib.util
    mod_name = 'ct_app_module'
    if mod_name in sys.modules:
        return sys.modules[mod_name]
    spec = importlib.util.spec_from_file_location(
        mod_name, os.path.join(_CT_DIR, 'app.py'))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod


def _df_to_tmp(df, name: str):
    """
    Persist a DataFrame to disk as Parquet (low RAM), fall back to CSV.
    Returns True on success.
    """
    path = _tmp_path(name)
    try:
        df.to_parquet(path, index=True, engine='pyarrow')
        return True
    except Exception:
        try:
            df.to_parquet(path, index=True, engine='fastparquet')
            return True
        except Exception:
            # Parquet not available — fall back to compressed CSV
            path = path.replace('.parquet', '.csv.gz')
            df.to_csv(path, index=True, compression='gzip')
            return True


def _df_from_tmp(name: str):
    """
    Load a DataFrame from the temp file written by _df_to_tmp.
    Returns None if file doesn't exist.
    """
    import pandas as pd
    path_pq  = _tmp_path(name)
    path_csv = path_pq.replace('.parquet', '.csv.gz')
    if os.path.exists(path_pq):
        try:
            return pd.read_parquet(path_pq, engine='pyarrow')
        except Exception:
            try:
                return pd.read_parquet(path_pq, engine='fastparquet')
            except Exception:
                pass
    if os.path.exists(path_csv):
        return pd.read_csv(path_csv, index_col=0, compression='gzip')
    return None


def _clear_all_tmp():
    """Delete all CT temp files (called on new /process to free disk)."""
    _del_tmp(_TF_CT_PARQUET, _TF_AN_PARQUET,
             _TF_WB_PARQUET, _TF_RJ_PARQUET,
             _TF_MAIN_XLSX,  _TF_SUMMARY_XLSX, _TF_PPT)
    # Also try CSV gz variants
    for name in [_TF_CT_PARQUET, _TF_AN_PARQUET,
                 _TF_WB_PARQUET, _TF_RJ_PARQUET]:
        try:
            os.remove(_tmp_path(name.replace('.parquet', '.csv.gz')))
        except FileNotFoundError:
            pass


def _make_date_suffix(start_dt, end_dt):
    if start_dt is None:
        return ''
    s_day = start_dt.date()
    e_day = end_dt.date() if end_dt else None
    if e_day is None or s_day == e_day:
        return start_dt.strftime('%Y%m%d')
    return f"{start_dt.strftime('%Y%m%d')}&{end_dt.strftime('%Y%m%d')}"


# ── Auth guard ────────────────────────────────────────────────────────────────
def _require_login():
    if 'username' not in session:
        return redirect(url_for('login_page'))
    return None


# ── Page ──────────────────────────────────────────────────────────────────────
@ct.route('/')
def ct_page():
    guard = _require_login()
    if guard:
        return guard
    return render_template_string(CT_PAGE_HTML)


# ── /ct/process — parse files, detect anomalies, persist DFs to disk ──────────
@ct.route('/process', methods=['POST', 'OPTIONS'])
def ct_process():
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard:
        return guard
    global _ct_store
    try:
        eng = _get_ct_engine()

        ct_files    = request.files.getlist('ct_files')
        online_file = request.files.get('online_file')
        from_s      = request.form.get('from_dt', '')
        to_s        = request.form.get('to_dt', '')
        reason      = request.form.get('reason', '')

        thresh_queue    = float(request.form.get('thresh_queue',    0))
        thresh_lag      = float(request.form.get('thresh_lag',      0))
        thresh_duration = float(request.form.get('thresh_duration', 120))
        # Lower-limit anomaly threshold: duration=0 OR duration < min_duration → Anomaly
        min_duration    = float(request.form.get('min_duration',    5))

        if not ct_files or ct_files[0].filename == '':
            return jsonify({'error': 'No Cycle Time files uploaded'}), 400
        if not online_file or online_file.filename == '':
            return jsonify({'error': 'No Online Data file uploaded'}), 400

        try:
            start_dt = datetime.fromisoformat(from_s) if from_s else None
            end_dt   = datetime.fromisoformat(to_s)   if to_s   else None
        except ValueError:
            start_dt = end_dt = None

        # Read all file bytes first, then release Flask file objects
        ct_ios = [io.BytesIO(f.read()) for f in ct_files]
        on_io  = io.BytesIO(online_file.read())
        # Release upload references so GC can reclaim them
        ct_files    = None
        online_file = None
        _gc()

        # ── Step A: Parse + split CT data ─────────────────────────────────────
        df_ct, df_an, fl, exc, ah = eng.prepare_ct_data(
            ct_ios, start_dt, end_dt,
            queue_minutes=thresh_queue,
            lag_minutes=thresh_lag,
            duration_threshold=thresh_duration,
            min_duration_threshold=min_duration,
        )
        ct_records = len(df_ct)
        ct_ios = None   # release upload bytes
        _gc()

        # ── Step B: Parse Online/WB data ──────────────────────────────────────
        df_wb, df_rj, nw = eng.prepare_online_data(on_io, start_dt, end_dt)
        wb_records = len(df_wb)
        on_io = None    # release upload bytes
        eng   = None    # release engine reference
        _gc()

        # ── Step C: Delete old temp files before writing new ones ─────────────
        _clear_all_tmp()

        # ── Step D: Persist DataFrames to disk, then free RAM ─────────────────
        _df_to_tmp(df_ct, _TF_CT_PARQUET);  df_ct = None;  _gc()
        _df_to_tmp(df_an, _TF_AN_PARQUET);  df_an = None;  _gc()
        _df_to_tmp(df_wb, _TF_WB_PARQUET);  df_wb = None;  _gc()
        _df_to_tmp(df_rj, _TF_RJ_PARQUET);  df_rj = None;  _gc()

        # Update lightweight store
        _ct_store.update({
            'ct_records':       ct_records,
            'wb_records':       wb_records,
            'exceedances':      exc,
            'applicable_hours': ah,
            'net_weight':       nw,
            'failure_list':     fl,   # list of 'fail'/'ok' — small
            'start_dt':         start_dt,
            'end_dt':           end_dt,
            'reason':           reason,
            'processed':        True,
            'main_done':        False,
            'sum_done':         False,
            'ppt_done':         False,
        })
        # ── Persist metadata to disk so it survives gunicorn worker restarts ──
        # (--max-requests recycles the worker process, wiping Python globals.
        #  The parquet files already survive on disk; now the metadata does too.)
        _save_meta()

        return jsonify({
            'ok':              True,
            'ct_records':      ct_records,
            'wb_records':      wb_records,
            'exceedances':     exc,
            'applicable_hours': ah,
            'net_weight':      nw,
        })

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


# ── /ct/build/main — 7-sheet Excel, reads DFs from disk ───────────────────────

# ── Background job state for /ct/build/main ────────────────────────────────────
# Render's HTTP proxy times out after ~30s even though gunicorn allows 900s.
# Solution: spawn build_full_report in a background thread, return immediately,
# let JS poll /ct/build/main/status every 5s until done.
import threading as _threading
_main_job = {
    'status': 'idle',   # idle | running | done | error
    'size_kb': 0,
    'error': '',
    'lock': _threading.Lock(),
}


def _run_build_main():
    """Background thread: builds the 7-sheet Excel report."""
    global _ct_store
    try:
        with _main_job['lock']:
            _main_job['status'] = 'running'
            _main_job['error']  = ''

        df_ct = _df_from_tmp(_TF_CT_PARQUET)
        df_an = _df_from_tmp(_TF_AN_PARQUET)
        df_wb = _df_from_tmp(_TF_WB_PARQUET)
        df_rj = _df_from_tmp(_TF_RJ_PARQUET)

        if df_ct is None:
            with _main_job['lock']:
                _main_job['status'] = 'error'
                _main_job['error']  = 'Processed data not found — please re-upload files'
            return

        eng  = _get_ct_engine()
        xlsx = eng.build_main_report(
            df_ct, df_an, df_wb, df_rj,
            _ct_store['failure_list'],
            _ct_store['start_dt'],
            _ct_store['end_dt'],
        )

        df_ct = df_an = df_wb = df_rj = eng = None
        _gc()

        _write_tmp(_TF_MAIN_XLSX, xlsx)
        size_kb = len(xlsx) // 1024
        xlsx = None
        _gc()

        _ct_store['main_done'] = True
        _save_meta()

        with _main_job['lock']:
            _main_job['status']  = 'done'
            _main_job['size_kb'] = size_kb

    except Exception as e:
        with _main_job['lock']:
            _main_job['status'] = 'error'
            _main_job['error']  = str(e)
        import traceback as _tb
        print(f"[CT build/main] BACKGROUND ERROR: {_tb.format_exc()}")


@ct.route('/build/main', methods=['POST', 'OPTIONS'])
def ct_build_main():
    """Start the 7-sheet Excel build in a background thread and return immediately.
    Render's HTTP proxy times out after ~30s — this avoids that by returning
    immediately with status='started' and letting the JS poll /ct/build/main/status.
    """
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard:
        return guard
    global _ct_store
    try:
        if not _load_meta():
            return jsonify({'error': 'Run /ct/process first'}), 400

        with _main_job['lock']:
            if _main_job['status'] == 'running':
                return jsonify({'status': 'running'}), 202
            _main_job['status']  = 'starting'
            _main_job['size_kb'] = 0
            _main_job['error']   = ''

        t = _threading.Thread(target=_run_build_main, daemon=True)
        t.start()
        return jsonify({'status': 'started'})

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@ct.route('/build/main/status', methods=['GET'])
def ct_build_main_status():
    """Poll this endpoint to check if the background Excel build has finished."""
    guard = _require_login()
    if guard:
        return guard
    with _main_job['lock']:
        status  = _main_job['status']
        size_kb = _main_job['size_kb']
        error   = _main_job['error']
    if status == 'done':
        return jsonify({'ok': True, 'status': 'done', 'size_kb': size_kb})
    elif status == 'error':
        return jsonify({'ok': False, 'status': 'error', 'error': error}), 500
    else:
        return jsonify({'status': status})  # 'starting' or 'running'




# ── /ct/build/summary — Hourly Trucks Excel ───────────────────────────────────
@ct.route('/build/summary', methods=['POST', 'OPTIONS'])
def ct_build_summary():
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard:
        return guard
    global _ct_store
    try:
        # Restore state from disk if worker was recycled since /ct/process ran
        if not _load_meta():
            return jsonify({'error': 'Run /ct/process first'}), 400

        # Only WB total needed for summary
        df_wb = _df_from_tmp(_TF_WB_PARQUET)
        if df_wb is None:
            return jsonify({'error': 'Processed data not found — please re-upload files'}), 400

        eng       = _get_ct_engine()
        demo_path = os.path.join(_CT_DIR, 'Demo.xlsx')
        xlsx = eng.build_summary_report(
            df_wb, demo_path,
            start_dt=_ct_store.get('start_dt'),
            end_dt=_ct_store.get('end_dt'),
        )

        df_wb = eng = None
        _gc()

        _write_tmp(_TF_SUMMARY_XLSX, xlsx)
        size_kb = len(xlsx) // 1024
        xlsx = None
        _gc()

        _ct_store['sum_done'] = True
        _save_meta()   # persist flag — survives worker restart before PPT build
        return jsonify({'ok': True, 'size_kb': size_kb})

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


# ── /ct/build/ppt — PowerPoint (requires summary to exist on disk) ─────────────
@ct.route('/build/ppt', methods=['POST', 'OPTIONS'])
def ct_build_ppt():
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard:
        return guard
    global _ct_store
    try:
        # Restore state from disk if worker was recycled since /ct/process ran
        if not _load_meta():
            return jsonify({'error': 'Run /ct/process first'}), 400
        if not _ct_store.get('sum_done'):
            return jsonify({'error': 'Build summary first (/ct/build/summary)'}), 400

        # Load only what PPT needs: CT clean + WB total + summary bytes
        df_ct = _df_from_tmp(_TF_CT_PARQUET)
        df_wb = _df_from_tmp(_TF_WB_PARQUET)
        summary_bytes = _read_tmp(_TF_SUMMARY_XLSX)

        if df_ct is None or summary_bytes is None:
            return jsonify({'error': 'Processed data not found — please re-upload files'}), 400

        eng = _get_ct_engine()
        ppt = eng.build_ppt_report(
            df_ct, df_wb,
            reason=_ct_store['reason'],
            exceedances=_ct_store['exceedances'],
            applicable_hours=_ct_store['applicable_hours'],
            summary_xlsx_bytes=summary_bytes,
            ppt_template_path=os.path.join(_CT_DIR, 'demo.pptx'),
        )

        df_ct = df_wb = summary_bytes = eng = None
        _gc()

        _write_tmp(_TF_PPT, ppt)
        size_kb = len(ppt) // 1024
        ppt = None
        _gc()

        _ct_store['ppt_done'] = True
        return jsonify({'ok': True, 'size_kb': size_kb})

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


# ── /ct/download/<kind> — stream from disk, no in-memory buffering ────────────
@ct.route('/download/<kind>', methods=['GET'])
def ct_download(kind):
    guard = _require_login()
    if guard:
        return guard

    _sd = _ct_store.get('start_dt')
    _ed = _ct_store.get('end_dt')
    _suffix   = _make_date_suffix(_sd, _ed)
    _ct_name  = f"Cycle Time - {_suffix}"  if _suffix else "Cycle Time"
    _hrq_name = f"APSG-Hourly Truck Quantity {_suffix}" if _suffix else "APSG-Hourly Truck Quantity"

    MAP = {
        'main':    (_TF_MAIN_XLSX,    f'{_ct_name}.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'summary': (_TF_SUMMARY_XLSX, f'{_hrq_name}.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'ppt':     (_TF_PPT,          f'{_ct_name}.pptx',
                    'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
    }
    if kind not in MAP:
        return jsonify({'error': 'Unknown type'}), 404

    tf_name, fname, mime = MAP[kind]
    data = _read_tmp(tf_name)
    if data is None:
        return jsonify({'error': f'{kind} not generated yet — please build first'}), 404

    resp = make_response(data)
    data = None   # release immediately after handing to Flask
    _gc()

    resp.headers['Content-Type'] = mime
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp

# ══════════════════════════════════════════════════════════════════════════════
#  CT PAGE HTML — follows ABSC Main theme exactly
# ══════════════════════════════════════════════════════════════════════════════
CT_PAGE_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cycle Time Report — APSG</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Poppins:wght@600;700;800&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --indigo:#6366F1;--indigo-l:#818CF8;--cyan:#22D3EE;
  --green:#10B981;--amber:#F59E0B;--red:#F87171;
  --text:#E8EEF8;--muted:#64748B;--border:rgba(99,102,241,0.15);
}
html,body{
  background-image:url('/static/bg.jpg') !important;
  background-size:cover !important;background-position:center !important;
  background-attachment:fixed !important;background-color:#08101E !important;
  font-family:'Inter',system-ui,sans-serif;min-height:100vh;color:var(--text);
  font-size:15px;font-weight:500;line-height:1.55;
}
body::before{content:'';position:fixed;inset:0;z-index:0;
  background:rgba(3,7,18,0.45);pointer-events:none;}
body>*{position:relative;z-index:1;}

/* ── Top bar — identical to ABSC Main ── */
.top-bar{position:sticky;top:0;z-index:200;height:40px;display:flex;align-items:center;
  padding:0 1.2rem;gap:.75rem;
  background:rgba(6,10,28,0.55) !important;backdrop-filter:blur(18px);
  border-bottom:1px solid rgba(255,255,255,0.08);}
.top-mini-brand{font-size:.75rem;font-weight:800;color:#6366F1;letter-spacing:.04em;white-space:nowrap;}
.top-sep{width:1px;height:18px;background:rgba(99,102,241,.18);}
.top-page-label{font-size:.78rem;font-weight:600;color:#FFFFFF;white-space:nowrap;}
.top-spacer{flex:1;}
.karthi-tag{font-size:.65rem;color:#6366F1;font-weight:700;white-space:nowrap;letter-spacing:.01em;}
.back-btn{background:rgba(99,102,241,0.18);border:1px solid rgba(99,102,241,0.35);
  border-radius:6px;padding:.22rem .7rem;font-size:.7rem;font-weight:600;
  color:#C5D5FF;text-decoration:none;white-space:nowrap;transition:all .2s;}
.back-btn:hover{background:rgba(99,102,241,0.32);}

/* ── Page body ── */
.page{padding:1rem 1.2rem 4rem;max-width:1100px;margin:0 auto;}

/* ── Cards ── */
.card{background:rgba(8,14,38,0.70);border:1px solid rgba(255,255,255,0.10);
  border-radius:14px;padding:1.2rem 1.4rem;margin-bottom:1rem;
  backdrop-filter:blur(16px);box-shadow:0 4px 32px rgba(0,0,0,0.35);}
.card-title{font-size:.9rem;font-weight:700;color:var(--indigo-l);margin-bottom:.9rem;
  display:flex;align-items:center;gap:.5rem;letter-spacing:.02em;}
.step-badge{display:inline-flex;align-items:center;justify-content:center;
  width:24px;height:24px;border-radius:50%;background:var(--indigo);
  color:#fff;font-size:.72rem;font-weight:800;flex-shrink:0;}

/* ── Side-by-side upload grid ── */
.upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
@media(max-width:600px){.upload-grid{grid-template-columns:1fr;}}
.upload-col-label{font-size:.68rem;font-weight:700;color:var(--muted);
  letter-spacing:.08em;text-transform:uppercase;margin-bottom:.4rem;}

/* ── Date dropdowns ── */
.date-sel{width:100%;padding:.72rem .9rem;
  background:rgba(5,9,28,.85);border:1.5px solid rgba(99,102,241,.45);
  border-radius:10px;color:#EEF3FF;font-size:15px;font-weight:600;
  font-family:'Inter',sans-serif;cursor:pointer;appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%236366F1' stroke-width='2' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right .9rem center;padding-right:2.2rem;}
.date-sel:focus{outline:none;border-color:var(--indigo);box-shadow:0 0 0 3px rgba(99,102,241,.2);}
.date-sel option{background:#0A0F2E;color:#EEF3FF;font-size:15px;}
.date-sel:disabled{opacity:.45;cursor:not-allowed;}

/* ── Upload zones ── */
.upload-zone{border:2px dashed rgba(99,102,241,0.55);border-radius:11px;
  padding:1.4rem;text-align:center;cursor:pointer;
  background:rgba(6,12,34,0.55);transition:all .2s;position:relative;}
.upload-zone:hover,.upload-zone.dragover{border-color:var(--indigo);background:rgba(99,102,241,.08);}
.upload-zone.ok{border-color:var(--green);background:rgba(16,185,129,.06);border-style:solid;}
.upload-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}
.file-list{display:flex;flex-wrap:wrap;gap:.35rem;margin-top:.6rem;}
.file-tag{background:rgba(99,102,241,.12);border:1px solid rgba(99,102,241,.25);
  color:#818CF8;padding:.2rem .6rem;border-radius:20px;font-size:.72rem;
  display:flex;align-items:center;gap:.3rem;}
.file-tag button{background:none;border:none;color:#F87171;cursor:pointer;font-size:.85rem;line-height:1;}

/* ── Anomaly config ── */
.anom-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
@media(max-width:800px){.anom-grid{grid-template-columns:repeat(2,1fr);}}
@media(max-width:500px){.anom-grid{grid-template-columns:1fr;}}
.anom-field{background:rgba(8,20,50,.6);border:1.5px solid rgba(99,102,241,.25);
  border-radius:10px;padding:12px 14px;}
.anom-field label{display:block;font-size:12px;font-weight:700;color:var(--indigo-l);
  margin-bottom:4px;letter-spacing:.3px;}
.anom-sub{font-size:11px;color:rgba(190,210,255,.55);margin-bottom:8px;line-height:1.4;}
.anom-input-wrap{display:flex;align-items:center;gap:8px;}
.anom-input{flex:1;border:1.5px solid rgba(99,102,241,.40);border-radius:7px;
  padding:10px 10px;font-size:15px;font-weight:700;color:#EEF3FF;
  background:rgba(5,9,28,.80);text-align:center;}
.anom-input:focus{outline:none;border-color:var(--indigo);box-shadow:0 0 0 3px rgba(99,102,241,.18);}
.anom-unit{font-size:13px;font-weight:700;color:var(--muted);white-space:nowrap;}
.anom-note{font-size:11px;color:#805ad5;margin-top:6px;font-weight:600;}

/* ── Reason dropdown ── */
select.reason-sel{width:100%;padding:.72rem .9rem;
  background:rgba(5,9,28,.80);border:1.5px solid rgba(99,102,241,.40);
  border-radius:10px;color:#EEF3FF;font-size:15px;font-weight:500;font-family:'Inter',sans-serif;}
select.reason-sel:focus{outline:none;border-color:var(--indigo);}

/* ── Date inputs ── */
.date-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
@media(max-width:500px){.date-row{grid-template-columns:1fr;}}
.date-field label{display:block;font-size:.68rem;font-weight:700;color:var(--muted);
  letter-spacing:.08em;text-transform:uppercase;margin-bottom:.4rem;}
input[type=datetime-local]{width:100%;padding:.65rem .9rem;
  background:rgba(5,9,28,.80);border:1.5px solid rgba(99,102,241,.40);
  border-radius:10px;color:#EEF3FF;font-size:14px;font-family:'Inter',sans-serif;}
input[type=datetime-local]:focus{outline:none;border-color:var(--indigo);}

/* ── Generate button ── */
.btn-gen{width:100%;padding:15px;border:none;border-radius:10px;
  background:linear-gradient(135deg,#2b6cb0,#1a365d);color:#fff;
  font-size:15px;font-weight:700;cursor:pointer;display:flex;
  align-items:center;justify-content:center;gap:10px;transition:all .2s;}
.btn-gen:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 8px 24px rgba(43,108,176,.4);}
.btn-gen:disabled{opacity:.55;cursor:not-allowed;transform:none;}

/* ── Progress cards (step indicators) ── */
.steps-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:.75rem;margin:.75rem 0;}
@media(max-width:640px){.steps-grid{grid-template-columns:1fr 1fr;}}
.step-item{border-radius:10px;padding:.85rem .9rem;text-align:center;border:1px solid rgba(255,255,255,.08);
  background:rgba(8,14,38,.6);transition:all .3s;}
.step-item.active{border-color:var(--amber);background:rgba(245,158,11,.06);}
.step-item.done{border-color:var(--green);background:rgba(16,185,129,.07);}
.step-item.error{border-color:var(--red);background:rgba(248,113,113,.06);}
.step-icon{font-size:1.4rem;margin-bottom:.35rem;}
.step-name{font-size:.78rem;font-weight:700;color:#EEF3FF;margin-bottom:.2rem;}
.step-stat{font-size:.72rem;color:var(--muted);}
.step-item.done .step-stat{color:var(--green);}
.step-item.error .step-stat{color:var(--red);}

/* ── Progress bar ── */
.prog-wrap{margin:.6rem 0;}
.prog-track{height:5px;background:rgba(99,102,241,.12);border-radius:3px;overflow:hidden;}
.prog-fill{height:100%;border-radius:3px;background:linear-gradient(90deg,#4338CA,#6366F1);
  width:0%;transition:width .4s;}
.prog-label{font-size:.75rem;font-weight:600;color:var(--muted);margin-top:.3rem;text-align:right;}

/* ── Percentage ticker ── */
.pct-ticker{display:none;align-items:center;gap:.65rem;margin:.5rem 0;padding:.5rem .9rem;
  background:rgba(99,102,241,.07);border:1px solid rgba(99,102,241,.18);border-radius:9px;}
.pct-num{font-size:1.55rem;font-weight:900;color:#818CF8;min-width:3.2rem;text-align:right;
  font-variant-numeric:tabular-nums;letter-spacing:-.02em;}
.pct-pct{font-size:.85rem;font-weight:700;color:#6366F1;padding-top:.25rem;}
.pct-msg{font-size:.78rem;font-weight:600;color:#94A3B8;flex:1;}
.pct-done{color:#34D399 !important;}

/* ── Log terminal ── */
.log-box{background:rgba(2,5,15,.85);border:1px solid rgba(99,102,241,.15);
  border-radius:8px;padding:.75rem 1rem;font-family:'Courier New',monospace;
  font-size:.78rem;line-height:1.8;max-height:180px;overflow-y:auto;margin:.5rem 0;}
.log-ok{color:#4ADE80;font-weight:600;} .log-err{color:#F87171;font-weight:600;} .log-norm{color:#94A3B8;}

/* ── Download section ── */
.dl-section{display:none;}
.success-banner{background:rgba(16,185,129,.09);border:1px solid rgba(16,185,129,.22);
  border-radius:10px;padding:.75rem 1rem;margin-bottom:1rem;
  color:#34D399;font-size:.82rem;font-weight:600;}
.btn-download-all{width:100%;padding:15px;border:none;border-radius:10px;
  background:linear-gradient(135deg,#276749,#48bb78);color:#fff;
  font-size:15px;font-weight:700;cursor:pointer;display:flex;
  align-items:center;justify-content:center;gap:10px;transition:all .2s;margin-bottom:.75rem;}
.btn-download-all:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(72,187,120,.4);}
.stats-row{display:flex;gap:.65rem;flex-wrap:wrap;margin-top:.75rem;}
.stat{flex:1;min-width:80px;background:rgba(99,102,241,.08);border:1px solid rgba(99,102,241,.15);
  border-radius:9px;padding:.75rem;text-align:center;}
.stat-val{font-size:1.5rem;font-weight:800;color:#818CF8;}
.stat-lbl{font-size:.62rem;color:var(--muted);margin-top:.15rem;}

/* ── Loading section ── */
.loading-section{display:none;}

/* ── Shimmer ── */
@keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
.shimmer-bar{height:4px;background:linear-gradient(90deg,
  rgba(99,102,241,.1) 0%,rgba(99,102,241,.5) 50%,rgba(99,102,241,.1) 100%);
  background-size:200% 100%;animation:shimmer 1.5s infinite;border-radius:2px;margin:.5rem 0;}

/* ── Spinner ── */
@keyframes spin{to{transform:rotate(360deg)}}
.spin{display:inline-block;animation:spin 1s linear infinite;}

/* Fixed dev footer */
.apsg-footer{position:fixed;bottom:0;left:0;right:0;z-index:9999;
  text-align:center;padding:.28rem 1rem;
  background:rgba(4,7,20,.70);backdrop-filter:blur(8px);
  border-top:1px solid rgba(255,255,255,.07);
  font-size:11px;font-weight:600;color:rgba(200,220,255,.70);
  letter-spacing:.06em;pointer-events:none;user-select:none;}
</style>
</head>
<body>

<!-- Top bar — matches ABSC Main exactly -->
<div class="top-bar">
  <span class="top-mini-brand">APSG</span>
  <div class="top-sep"></div>
  <span class="top-page-label">⏱ Cycle Time Report</span>
  <div class="top-spacer"></div>
  <span class="karthi-tag">✦ Karthi</span>
  <a href="/" class="back-btn">← Dashboard</a>
</div>

<div class="page">

  <!-- STEP 1: Upload Files — CT left, Online right -->
  <div class="card">
    <div class="card-title"><span class="step-badge">1</span> Upload Files</div>
    <div class="upload-grid">

      <!-- LEFT: Cycle Time Files (primary — drives date dropdowns) -->
      <div>
        <div class="upload-col-label">📊 Cycle Time Files <span style="font-size:.62rem;color:var(--muted);font-weight:400">(one or more)</span></div>
        <div class="upload-zone" id="ctZone">
          <input type="file" id="ctFiles" multiple accept=".xlsx,.xls,.csv" onchange="addCTFiles(this.files)">
          <div style="font-size:1.5rem;margin-bottom:.3rem">📊</div>
          <div style="font-size:.82rem;font-weight:600">Drop CT files or click</div>
          <div style="font-size:.7rem;color:var(--muted);margin-top:.2rem">.xlsx · .xls · .csv — multiple OK</div>
        </div>
        <div class="file-list" id="ctFileList"></div>
      </div>

      <!-- RIGHT: Online Data File -->
      <div>
        <div class="upload-col-label">📋 Online Data File</div>
        <div class="upload-zone" id="onZone">
          <input type="file" id="onFile" accept=".xlsx,.xls,.csv" onchange="setOnFile(this.files[0])">
          <div style="font-size:1.5rem;margin-bottom:.3rem">📋</div>
          <div style="font-size:.82rem;font-weight:600">Drop Online file or click</div>
          <div style="font-size:.7rem;color:var(--muted);margin-top:.2rem">.xlsx · .xls · .csv</div>
        </div>
        <div id="onFileName" style="font-size:.73rem;color:var(--green);margin-top:.4rem;font-weight:600;min-height:1rem;"></div>
      </div>

    </div>
  </div>

  <!-- STEP 2: Date Range -->
  <div class="card">
    <div class="card-title"><span class="step-badge">2</span> Date &amp; Time Range
      <span id="date-auto-badge" style="display:none;margin-left:8px;font-size:.65rem;font-weight:700;background:rgba(16,185,129,.15);color:#34D399;border:1px solid rgba(16,185,129,.25);border-radius:5px;padding:.15rem .55rem;">✦ Dates loaded from file</span>
    </div>

    <!-- Info bar: times are always fixed -->
    <div style="display:flex;align-items:center;gap:.6rem;margin-bottom:.85rem;padding:.6rem .9rem;background:rgba(99,102,241,.07);border:1px solid rgba(99,102,241,.2);border-radius:9px;">
      <span style="font-size:1rem;">⏰</span>
      <span style="font-size:.82rem;font-weight:600;color:#C5D5FF;">
        Time is always &nbsp;<span style="color:var(--indigo-l);font-weight:700;">From → 07:00 AM &nbsp;·&nbsp; To → 05:00 AM</span>
        &nbsp;<span style="color:var(--muted);font-weight:400;font-size:.73rem;">— upload a file to select available dates below</span>
      </span>
    </div>

    <div class="date-row" id="date-inputs">
      <div class="date-field">
        <label style="font-size:.75rem;font-weight:700;color:var(--muted);letter-spacing:.06em;text-transform:uppercase;display:block;margin-bottom:.45rem;">From Date</label>
        <select id="from-dt-sel" class="date-sel" onchange="_syncHiddenDates()">
          <option value="">— Upload a file first —</option>
        </select>
        <input type="hidden" id="from-dt" value="">
      </div>
      <div class="date-field">
        <label style="font-size:.75rem;font-weight:700;color:var(--muted);letter-spacing:.06em;text-transform:uppercase;display:block;margin-bottom:.45rem;">To Date</label>
        <select id="to-dt-sel" class="date-sel" onchange="_syncHiddenDates()">
          <option value="">— Upload a file first —</option>
        </select>
        <input type="hidden" id="to-dt" value="">
      </div>
    </div>

    <div style="margin-top:.85rem;">
      <div style="font-size:.75rem;font-weight:700;color:var(--muted);letter-spacing:.08em;text-transform:uppercase;margin-bottom:.45rem;">Queue Condition</div>
      <select id="reason" class="reason-sel">
        <option value="There is a queue condition due to hourly loads are higher than expected.">Queue condition — hourly loads higher than expected</option>
        <option value="There is a queue condition due to heavy rain.">Queue condition — heavy rain</option>
        <option value="There is no queue condition.">No queue condition</option>
        <option value="There is a queue condition due to weighbridge breakdown.">Queue condition — weighbridge breakdown</option>
      </select>
    </div>
  </div>

  <!-- STEP 2b: Anomaly Thresholds -->
  <div class="card">
    <div class="card-title"><span class="step-badge" style="background:#805ad5">⚙</span> Anomaly Detection Thresholds <span style="font-size:10px;font-weight:400;color:var(--muted);margin-left:6px">rows exceeding any threshold → Anomaly sheet</span></div>
    <div class="anom-grid">
      <div class="anom-field">
        <label>⏱ Date Time Arrival → WB In Time</label>
        <div class="anom-sub">Pre-weighbridge queue time</div>
        <div class="anom-input-wrap">
          <input type="number" id="thresh-queue" class="anom-input" value="0" min="0" max="9999" step="1">
          <span class="anom-unit">min</span>
        </div>
        <div class="anom-note">0 = skip this check entirely</div>
      </div>
      <div class="anom-field">
        <label>⏱ WB Out Time → Date Time Exit</label>
        <div class="anom-sub">Post-weighbridge lag time</div>
        <div class="anom-input-wrap">
          <input type="number" id="thresh-lag" class="anom-input" value="0" min="0" max="9999" step="1">
          <span class="anom-unit">min</span>
        </div>
        <div class="anom-note">0 = skip this check entirely</div>
      </div>
      <div class="anom-field">
        <label>⏱ Duration — Upper Limit</label>
        <div class="anom-sub">Duration column from uploaded CT Excel</div>
        <div class="anom-input-wrap">
          <input type="number" id="thresh-duration" class="anom-input" value="70" min="1" max="9999" step="1">
          <span class="anom-unit">min</span>
        </div>
        <div class="anom-note">Rows where Duration &gt; this value → Anomaly</div>
      </div>
      <div class="anom-field">
        <label>⏱ Duration — Lower Limit</label>
        <div class="anom-sub">Auto-Anomaly if duration is too short</div>
        <div class="anom-input-wrap">
          <input type="number" id="thresh-min-duration" class="anom-input" value="5" min="0" max="60" step="1">
          <span class="anom-unit">min</span>
        </div>
        <div class="anom-note">Duration = 0 OR &lt; this value → always Anomaly</div>
      </div>
    </div>
  </div>

  <!-- STEP 3: Generate -->
  <div class="card">
    <div class="card-title"><span class="step-badge">3</span> Generate Report</div>
    <button class="btn-gen" id="btn-gen" onclick="runGenerate()">🚀 &nbsp; Generate Report</button>
  </div>

  <!-- Loading / progress -->
  <div class="loading-section" id="loading-section">
    <div class="card">
      <div class="shimmer-bar"></div>
      <div class="steps-grid" id="steps-grid">
        <div class="step-item" id="scard-0"><div class="step-icon" id="sicon-0">📁</div><div class="step-name">Processing Files</div><div class="step-stat" id="sstat-0">Waiting…</div></div>
        <div class="step-item" id="scard-1"><div class="step-icon" id="sicon-1">📊</div><div class="step-name">Cycle Time Report</div><div class="step-stat" id="sstat-1">Waiting…</div></div>
        <div class="step-item" id="scard-2"><div class="step-icon" id="sicon-2">📋</div><div class="step-name">Hourly Track</div><div class="step-stat" id="sstat-2">Waiting…</div></div>
        <div class="step-item" id="scard-3"><div class="step-icon" id="sicon-3">📑</div><div class="step-name">PowerPoint</div><div class="step-stat" id="sstat-3">Waiting…</div></div>
      </div>
      <div class="prog-wrap">
        <div class="prog-track"><div class="prog-fill" id="prog-fill"></div></div>
        <div class="prog-label" id="prog-label"></div>
      </div>
      <!-- Lightweight % progress ticker — no real-time server polling, pure JS timer -->
      <div class="pct-ticker" id="pct-ticker">
        <span class="pct-num" id="pct-num">0</span>
        <span class="pct-pct">%</span>
        <span class="pct-msg" id="pct-msg">Starting…</span>
      </div>
      <div class="log-box" id="log"></div>
    </div>
  </div>

  <!-- Download section -->
  <div class="dl-section" id="dl-section">
    <div class="card">
      <div class="card-title"><span class="step-badge">4</span> Download Your Reports</div>
      <div class="success-banner">✅ &nbsp; All 3 reports generated successfully and ready for download!</div>
      <button class="btn-download-all" onclick="downloadAll()">⬇ &nbsp; Download Report</button>
      <div class="stats-row" id="stats"></div>
    </div>
  </div>

</div><!-- /page -->

<div class="apsg-footer">✦ Internal Reporting Platform — APSG Staging Ground &nbsp;·&nbsp; Developed by Karthik</div>

<script>
// ── SheetJS CDN (lazy-loaded on first upload) ─────────────────────────────────
let _XLSX = null;
async function _loadXLSX() {
  if (_XLSX) return _XLSX;
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = () => { _XLSX = window.XLSX; resolve(_XLSX); };
    s.onerror = () => reject(new Error('SheetJS failed to load'));
    document.head.appendChild(s);
  });
}

// ── File state ────────────────────────────────────────────────────────────────
const _files = { ct: [], on: null };

// ── Extract unique calendar dates from uploaded file ─────────────────────────
// Priority: WB In Time column → Date Time Arrival column → all cells scan.
// Returns sorted array of 'YYYY-MM-DD' strings.
async function _extractDatesFromFile(file) {
  try {
    const XLSX = await _loadXLSX();
    const buf  = await file.arrayBuffer();
    // Read with cellDates:true AND raw:true so we get both Date objects and raw values
    const wb   = XLSX.read(buf, { type: 'array', cellDates: true, raw: true });
    const ws   = wb.Sheets[wb.SheetNames[0]];

    // Get rows with raw values (Date objects, numbers, strings)
    const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    // Also get formatted string rows for string-based date parsing
    const strRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });

    if (rawRows.length < 2) return [];

    const fmt = d => {
      const y = d.getFullYear(), m = d.getMonth()+1, day = d.getDate();
      return `${y}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    };

    // Find the target column index from the header row
    const header = rawRows[0].map(h => String(h || '').trim().toLowerCase());
    let colIdx = -1;

    // Try WB In Time first
    colIdx = header.findIndex(h =>
      h === 'wb in time' || h === 'wbintime' || h === 'wb_in_time' ||
      (h.includes('wb') && (h.includes('in time') || h === 'wb in'))
    );
    // Try Date Time Arrival as first fallback
    if (colIdx < 0) {
      colIdx = header.findIndex(h =>
        h.includes('arrival') || h === 'date time arrival' || h === 'datetime arrival'
      );
    }
    // Try any column with "date" or "time" in the name
    if (colIdx < 0) {
      colIdx = header.findIndex(h => h.includes('date') || h.includes('time'));
    }

    const dateSet = new Set();

    // ── Pass 1: scan the target column using raw values ──────────────────────
    if (colIdx >= 0) {
      for (let r = 1; r < rawRows.length; r++) {
        const raw = rawRows[r][colIdx];
        const str = strRows[r] ? strRows[r][colIdx] : null;
        const d = _parseCell(XLSX, raw, str);
        if (d) dateSet.add(fmt(d));
      }
    }

    // ── Pass 2: if still empty, scan ALL columns ──────────────────────────────
    if (dateSet.size === 0) {
      for (let r = 1; r < Math.min(rawRows.length, 3000); r++) {
        for (let c = 0; c < (rawRows[r] || []).length; c++) {
          const raw = rawRows[r][c];
          const str = strRows[r] ? strRows[r][c] : null;
          const d = _parseCell(XLSX, raw, str);
          if (d) dateSet.add(fmt(d));
        }
      }
    }

    return [...dateSet].sort();
  } catch(e) {
    console.error('Date extraction failed:', e);
    return [];
  }
}

// ── Parse a single cell value into a JS Date (or null) ───────────────────────
function _parseCell(XLSX, raw, str) {
  // 1. Already a JS Date object
  if (raw instanceof Date && !isNaN(raw) && raw.getFullYear() > 2000) return raw;

  // 2. Excel serial number (date: 40000–60000, datetime has decimal)
  if (typeof raw === 'number' && raw > 40000 && raw < 60000) {
    try {
      const info = XLSX.SSF.parse_date_code(Math.floor(raw));
      if (info && info.y > 2000) return new Date(info.y, info.m - 1, info.d);
    } catch(e) {}
  }

  // 3. String parsing — try the formatted value first, then the raw string
  const candidates = [str, raw].filter(v => typeof v === 'string' && v.trim());
  for (const s of candidates) {
    const t = s.trim();
    // ISO: 2026-04-21 or 2026-04-21T07:00:00
    const iso = t.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})/);
    if (iso) {
      const d = new Date(+iso[1], +iso[2]-1, +iso[3]);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
    // Named month: 21-Apr-2026 or 21 Apr 2026
    const named = t.match(/^(\d{1,2})[\s\-](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\s\-](\d{4})/i);
    if (named) {
      const mo = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
      const d = new Date(+named[3], mo[named[2].toLowerCase()], +named[1]);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
    // DMY: 21/04/2026 or 21-04-2026
    const dmy = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      const d = new Date(+dmy[3], +dmy[2]-1, +dmy[1]);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
    // MDY fallback: 04/21/2026
    const mdy = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (mdy) {
      const d = new Date(t);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
  }
  return null;
}

// ── Populate From/To dropdowns from date array ────────────────────────────────
function _populateDateDropdowns(dates, prevFrom, prevTo) {
  const fromSel = document.getElementById('from-dt-sel');
  const toSel   = document.getElementById('to-dt-sel');
  if (!dates.length) return;

  function _label(ymd) {
    // Use noon to avoid any timezone-shift issues flipping the date
    const d = new Date(ymd + 'T12:00:00');
    return d.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric', weekday:'short' });
  }

  const opts = dates.map(d => `<option value="${d}">${_label(d)}</option>`).join('');
  fromSel.innerHTML = opts;
  toSel.innerHTML   = opts;

  // Restore previous selection if it still exists, else default earliest/latest
  fromSel.value = (prevFrom && dates.includes(prevFrom)) ? prevFrom : dates[0];
  toSel.value   = (prevTo   && dates.includes(prevTo))   ? prevTo   : dates[dates.length - 1];
  _syncHiddenDates();

  document.getElementById('date-auto-badge').style.display = 'inline-block';
}

// ── Sync hidden datetime-local values (always 07:00 From / 05:00 To) ─────────
function _syncHiddenDates() {
  const fromDate = document.getElementById('from-dt-sel').value;
  const toDate   = document.getElementById('to-dt-sel').value;
  document.getElementById('from-dt').value = fromDate ? fromDate + 'T07:00' : '';
  document.getElementById('to-dt').value   = toDate   ? toDate   + 'T05:00' : '';
}

// ── File upload handlers ──────────────────────────────────────────────────────

// Scan ALL current CT files, merge all unique dates, populate dropdowns.
// Preserves current From/To selections if the dates still exist after a change.
async function _refreshCTDates() {
  if (_files.ct.length === 0) {
    // Reset dropdowns
    const fromSel = document.getElementById('from-dt-sel');
    const toSel   = document.getElementById('to-dt-sel');
    fromSel.innerHTML = '<option value="">— Upload a file first —</option>';
    toSel.innerHTML   = '<option value="">— Upload a file first —</option>';
    document.getElementById('from-dt').value = '';
    document.getElementById('to-dt').value   = '';
    document.getElementById('date-auto-badge').style.display = 'none';
    return;
  }

  // Remember current selections before rebuilding
  const prevFrom = document.getElementById('from-dt-sel').value;
  const prevTo   = document.getElementById('to-dt-sel').value;

  // Extract dates from EVERY CT file and merge into one sorted unique set
  const allDatesSet = new Set();
  for (const file of _files.ct) {
    const dates = await _extractDatesFromFile(file);
    dates.forEach(d => allDatesSet.add(d));
  }
  const allDates = [...allDatesSet].sort();

  if (allDates.length === 0) return;

  _populateDateDropdowns(allDates, prevFrom, prevTo);
}

async function addCTFiles(files) {
  Array.from(files).forEach(f => {
    if (!_files.ct.find(x => x.name === f.name && x.size === f.size)) _files.ct.push(f);
  });
  renderCTList();
  document.getElementById('ctZone').classList.toggle('ok', _files.ct.length > 0);
  await _refreshCTDates();
}

function removeCT(i) {
  _files.ct.splice(i, 1);
  renderCTList();
  document.getElementById('ctZone').classList.toggle('ok', _files.ct.length > 0);
  _refreshCTDates();  // Re-scan remaining files
}

function renderCTList() {
  document.getElementById('ctFileList').innerHTML = _files.ct.map((f,i) =>
    `<div class="file-tag">📄 ${esc(f.name)} <button onclick="removeCT(${i})">×</button></div>`
  ).join('');
}

async function setOnFile(f) {
  _files.on = f;
  document.getElementById('onFileName').textContent = f ? '✓ ' + f.name : '';
  document.getElementById('onZone').classList.toggle('ok', !!f);
  // Use Online file dates only if no CT files have been uploaded yet
  if (f && _files.ct.length === 0 && !document.getElementById('from-dt').value) {
    const dates = await _extractDatesFromFile(f);
    if (dates.length) _populateDateDropdowns(dates, '', '');
  }
}

// Drag-drop
const ctZ = document.getElementById('ctZone');
ctZ.addEventListener('dragover',  e => { e.preventDefault(); ctZ.classList.add('dragover'); });
ctZ.addEventListener('dragleave', () => ctZ.classList.remove('dragover'));
ctZ.addEventListener('drop', e => { e.preventDefault(); ctZ.classList.remove('dragover'); addCTFiles(e.dataTransfer.files); });

const onZ = document.getElementById('onZone');
onZ.addEventListener('dragover',  e => { e.preventDefault(); onZ.classList.add('dragover'); });
onZ.addEventListener('dragleave', () => onZ.classList.remove('dragover'));
onZ.addEventListener('drop', e => { e.preventDefault(); onZ.classList.remove('dragover'); setOnFile(e.dataTransfer.files[0]); });

// ── Step / progress helpers ───────────────────────────────────────────────────
function setStep(i, state, stat) {
  document.getElementById('scard-'+i).className = 'step-item' + (state ? ' '+state : '');
  const s = document.getElementById('sstat-'+i);
  if (s) s.textContent = stat || '';
}
function setProgress(pct, label) {
  document.getElementById('prog-fill').style.width = pct + '%';
  document.getElementById('prog-label').textContent = label || '';
}
function addLog(msg, cls) {
  const lb = document.getElementById('log');
  const d  = document.createElement('div');
  d.className = 'log-' + (cls || 'norm');
  d.textContent = '[' + new Date().toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',second:'2-digit'}) + '] ' + msg;
  lb.appendChild(d);
  lb.scrollTop = lb.scrollHeight;
}

// ── Lightweight % percentage ticker ──────────────────────────────────────────
// Pure client-side timer — NO server polling, NO extra requests.
// Smoothly counts from a "from" value toward a "target" value, stopping just
// before it reaches target so it never shows 100% until the step is truly done.
// This keeps RAM/CPU on the server completely unaffected.
let _pctTimer  = null;   // interval handle
let _pctCur    = 0;      // current displayed percentage
let _pctTarget = 0;      // target to crawl toward
const _pctEl   = () => document.getElementById('pct-num');
const _pctMsg  = () => document.getElementById('pct-msg');
const _pctBox  = () => document.getElementById('pct-ticker');

function _pctShow() { _pctBox().style.display = 'flex'; }
function _pctHide() { _pctBox().style.display = 'none'; }

function _pctSet(target, msg) {
  _pctTarget = target;
  if (msg) _pctMsg().textContent = msg;
}

function _pctDone(msg) {
  // Jump immediately to 100% and show completion message
  if (_pctTimer) { clearInterval(_pctTimer); _pctTimer = null; }
  _pctCur = 100;
  const el = _pctEl();
  if (el) { el.textContent = '100'; el.classList.add('pct-done'); }
  if (msg) { const m = _pctMsg(); if(m){ m.textContent = msg; m.classList.add('pct-done'); } }
}

function _pctReset() {
  if (_pctTimer) { clearInterval(_pctTimer); _pctTimer = null; }
  _pctCur = 0; _pctTarget = 0;
  const el = _pctEl(); if(el){ el.textContent='0'; el.classList.remove('pct-done'); }
  const m = _pctMsg(); if(m){ m.textContent='Starting…'; m.classList.remove('pct-done'); }
  _pctHide();
}

function _pctStart() {
  _pctShow();
  // Tick every 600ms: advance current toward target by 1-2%, never exceeding target
  // On 0.1 CPU the steps are slow — a gentle crawl looks natural and reassuring
  _pctTimer = setInterval(() => {
    if (_pctCur < _pctTarget) {
      _pctCur = Math.min(_pctCur + 1, _pctTarget);
      const el = _pctEl(); if(el) el.textContent = _pctCur;
    }
  }, 600);
}

// ── Download all ──────────────────────────────────────────────────────────────
function downloadAll() {
  [['main','Full_Report.xlsx'],['summary','Hourly_Report.xlsx'],['ppt','Report.pptx']].forEach(([k,n],i) => {
    setTimeout(() => {
      const a = document.createElement('a');
      a.href = '/ct/download/' + k; a.download = n;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
    }, i * 700);
  });
}

// ── Main generate ─────────────────────────────────────────────────────────────
async function runGenerate() {
  if (!_files.ct.length) { alert('⚠️ Please upload at least one Cycle Time file.'); return; }
  if (!_files.on)        { alert('⚠️ Please upload the Online Data file.'); return; }

  const fromDt = document.getElementById('from-dt').value;
  const toDt   = document.getElementById('to-dt').value;
  if (!fromDt || !toDt) {
    alert('⚠️ No dates available yet — please upload a Cycle Time or Online file first so dates can be loaded.');
    return;
  }

  document.getElementById('btn-gen').disabled = true;
  document.getElementById('dl-section').style.display = 'none';
  document.getElementById('loading-section').style.display = 'block';
  document.getElementById('log').innerHTML = '';
  for (let i=0; i<4; i++) setStep(i,'','Waiting…');
  setProgress(0,'');
  _pctReset();   // reset % ticker

  const reason  = document.getElementById('reason').value;
  const threshQ = parseFloat(document.getElementById('thresh-queue').value)       || 0;
  const threshL = parseFloat(document.getElementById('thresh-lag').value)         || 0;
  const threshD = parseFloat(document.getElementById('thresh-duration').value)    || 70;
  const minDur  = parseFloat(document.getElementById('thresh-min-duration').value) || 5;

  const fd = new FormData();
  _files.ct.forEach(f => fd.append('ct_files', f));
  fd.append('online_file', _files.on);
  fd.append('from_dt', fromDt);
  fd.append('to_dt',   toDt);
  fd.append('reason',  reason);
  fd.append('thresh_queue',    threshQ);
  fd.append('thresh_lag',      threshL);
  fd.append('thresh_duration', threshD);
  fd.append('min_duration',    minDur);

  let j1;
  try {
    // ── Step 0: Process ─────────────────────────────────────────────────
    setStep(0, 'active', 'Processing…');
    setProgress(5, 'Reading and validating uploaded files…');
    _pctSet(18, 'Reading & validating uploaded files…'); _pctStart();
    const r1 = await fetch('/ct/process', {method:'POST', body:fd});
    try { j1 = await r1.json(); } catch(e) { j1 = {error:'HTTP '+r1.status}; }
    if (!r1.ok || j1.error) { setStep(0,'error','Failed'); addLog('❌ '+j1.error,'err'); done(); return; }
    setStep(0,'done',`CT:${j1.ct_records} WB:${j1.wb_records}`);
    addLog(`✅ Files processed — CT: ${j1.ct_records}, WB: ${j1.wb_records} records`,'ok');
    setProgress(20,'Building Cycle Time Report…');

    // ── Step 1: Main Excel — start background job then poll ─────────────────
    setStep(1,'active','Building…'); setProgress(25,'Building Cycle Time Report…');
    _pctSet(62, 'Building Cycle Time Report (7 sheets)…');

    // Kick off the background build (returns immediately with {status:'started'})
    const r2start = await fetch('/ct/build/main',{method:'POST'});
    let j2start;
    try { j2start = await r2start.json(); } catch(e) { j2start = {error:'Unexpected response'}; }
    if (!r2start.ok || j2start.error) {
      setStep(1,'error','Failed'); addLog('❌ '+(j2start.error||'Failed to start build'),'err'); done(); return;
    }

    // Poll /ct/build/main/status every 6 seconds until done or error
    let j2 = null;
    for (let attempt = 0; attempt < 120; attempt++) {  // max 120 × 6s = 12 min
      await new Promise(res => setTimeout(res, 6000));  // wait 6 seconds
      let r2p;
      try {
        r2p = await fetch('/ct/build/main/status', {method:'GET'});
        j2  = await r2p.json();
      } catch(e) {
        j2 = {status:'running'};   // network hiccup — keep polling
        continue;
      }
      if (j2.status === 'done')  break;
      if (j2.status === 'error') break;
      // still running — update percentage display
      if (_pctCur < 61) _pctSet(Math.min(_pctCur + 2, 61), 'Building Cycle Time Report (7 sheets)…');
    }

    if (!j2 || j2.status === 'error' || !j2.ok) {
      setStep(1,'error','Failed');
      addLog('❌ '+(j2?.error || 'CT Report build failed'),'err');
      done(); return;
    }
    setStep(1,'done',j2.size_kb+' KB'); addLog('✅ Cycle Time Report ready','ok');
    setProgress(65,'Building Hourly Track…');

    // ── Step 2: Summary ──────────────────────────────────────────────────
    setStep(2,'active','Building…'); setProgress(68,'Building Hourly Track…');
    _pctSet(78, 'Building Hourly Trucks Quantity Report…');
    const r3 = await fetch('/ct/build/summary',{method:'POST'});
    let j3; try{j3=await r3.json();}catch(e){j3={error:'Unexpected response'};}
    if(!r3.ok||j3.error){setStep(2,'error','Failed');addLog('❌ '+j3.error,'err');done();return;}
    setStep(2,'done',j3.size_kb+' KB'); addLog('✅ Hourly Track ready','ok');
    setProgress(80,'Generating PowerPoint…');

    // ── Step 3: PowerPoint ───────────────────────────────────────────────
    setStep(3,'active','Building…'); setProgress(82,'Generating PowerPoint…');
    _pctSet(98, 'Generating PowerPoint report…');
    const r4 = await fetch('/ct/build/ppt',{method:'POST'});
    let j4; try{j4=await r4.json();}catch(e){j4={error:'Unexpected response'};}
    if(!r4.ok||j4.error){setStep(3,'error','Failed');addLog('❌ '+j4.error,'err');done();return;}
    setStep(3,'done',j4.size_kb+' KB'); addLog('✅ PowerPoint ready','ok');
    setProgress(100,'✅ All reports complete!');
    _pctDone('✅ 100% — All reports complete! Download is ready.');

    addLog('🎉 All reports ready — click Download Report below.','ok');
    const s = j1;
    document.getElementById('stats').innerHTML = `
      <div class="stat"><div class="stat-val">${s.ct_records||0}</div><div class="stat-lbl">CT Records</div></div>
      <div class="stat"><div class="stat-val">${s.wb_records||0}</div><div class="stat-lbl">WB Records</div></div>
      <div class="stat"><div class="stat-val">${s.exceedances||0}</div><div class="stat-lbl">Exceedances</div></div>
      <div class="stat"><div class="stat-val">${s.applicable_hours||0}</div><div class="stat-lbl">Applicable Hours</div></div>`;
    document.getElementById('dl-section').style.display = 'block';
    document.getElementById('dl-section').scrollIntoView({behavior:'smooth'});

  } catch(e) {
    addLog('❌ Connection error: '+e.message,'err');
  }
  done();
}

function done() {
  document.getElementById('btn-gen').disabled = false;
  // Stop the % crawl timer on error or unexpected exit
  if (_pctTimer) { clearInterval(_pctTimer); _pctTimer = null; }
}
function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;'); }
</script>
</body>
</html>"""
