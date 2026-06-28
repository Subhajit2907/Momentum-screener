"""
Flask backend for Nifty Midcap 100 EMA Screener.
Endpoints:
  GET  /          → serves the web UI
  POST /run       → accepts CSV, runs screener, returns JSON results
  GET  /download  → returns the generated Excel file
"""

import io
import json
import os
import threading
import time
import uuid

from flask import Flask, jsonify, render_template, request, send_file, session

import screener as sc

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "nifty-screener-secret-2024")

# In-memory job store { job_id: { status, progress, result, excel_bytes } }
jobs = {}
jobs_lock = threading.Lock()


def run_job(job_id, file_bytes):
    """Runs screener in a background thread and stores result in jobs dict."""
    def update_progress(msg):
        with jobs_lock:
            jobs[job_id]["progress"] = msg

    try:
        with jobs_lock:
            jobs[job_id]["status"] = "running"

        result = sc.run_full_screen(file_bytes, progress_callback=update_progress)

        with jobs_lock:
            jobs[job_id]["status"]      = "done"
            jobs[job_id]["result"]      = {
                "top10":      result["top10"],
                "all_passed": result["all_passed"],
                "rejected":   result["rejected"],
                "stats":      result["stats"],
                "run_date":   result["run_date"],
                "has_excel":  result["excel_bytes"] is not None,
            }
            jobs[job_id]["excel_bytes"] = result["excel_bytes"]
            jobs[job_id]["progress"]    = "Complete"

    except Exception as e:
        with jobs_lock:
            jobs[job_id]["status"]   = "error"
            jobs[job_id]["progress"] = f"Error: {str(e)}"
            jobs[job_id]["error"]    = str(e)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/run", methods=["POST"])
def run():
    """Accept CSV upload, kick off background job, return job_id."""
    if "csv_file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["csv_file"]
    if not f.filename.endswith(".csv"):
        return jsonify({"error": "Please upload a .csv file"}), 400

    file_bytes = f.read()
    job_id = str(uuid.uuid4())

    with jobs_lock:
        jobs[job_id] = {
            "status":      "queued",
            "progress":    "Starting…",
            "result":      None,
            "excel_bytes": None,
            "error":       None,
            "created_at":  time.time(),
        }

    t = threading.Thread(target=run_job, args=(job_id, file_bytes), daemon=True)
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    """Poll job status and progress."""
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status":   job["status"],
        "progress": job["progress"],
        "result":   job["result"],
        "error":    job.get("error"),
    })


@app.route("/download/<job_id>")
def download(job_id):
    """Download the Excel file for a completed job."""
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Result not ready"}), 404
    if not job["excel_bytes"]:
        return jsonify({"error": "No Excel file generated"}), 404

    run_date = job["result"]["run_date"].replace(" ", "_")
    filename = f"Midcap_EMA_Screen_{run_date}.xlsx"

    return send_file(
        io.BytesIO(job["excel_bytes"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


# Clean up jobs older than 2 hours
def cleanup_old_jobs():
    while True:
        time.sleep(3600)
        cutoff = time.time() - 7200
        with jobs_lock:
            old = [k for k, v in jobs.items() if v["created_at"] < cutoff]
            for k in old:
                del jobs[k]

threading.Thread(target=cleanup_old_jobs, daemon=True).start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
