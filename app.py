import os
import time
import uuid
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

from report_engine import ReportGenerationError, generate_report


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {"csv"}
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "16"))
FILE_TTL_SECONDS = int(os.environ.get("FILE_TTL_SECONDS", "3600"))


app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key-in-production")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["UPLOAD_FOLDER"] = str(UPLOAD_DIR)
app.config["OUTPUT_FOLDER"] = str(OUTPUT_DIR)

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def cleanup_old_files():
    now = time.time()
    for folder in (UPLOAD_DIR, OUTPUT_DIR):
        for path in folder.iterdir():
            if path.is_file() and path.name != ".gitkeep":
                try:
                    if now - path.stat().st_mtime > FILE_TTL_SECONDS:
                        path.unlink()
                except OSError:
                    app.logger.warning("Could not delete temporary file: %s", path)


def save_upload(file_storage, prefix):
    original_name = secure_filename(file_storage.filename or "")
    if not original_name:
        raise ValueError("Please select both CSV files.")
    if not allowed_file(original_name):
        raise ValueError("Only .csv files are allowed.")

    unique_name = f"{prefix}_{uuid.uuid4().hex}_{original_name}"
    destination = UPLOAD_DIR / unique_name
    file_storage.save(destination)
    return destination


def output_path_for(report_id, file_type):
    if file_type not in {"pdf", "excel"}:
        return None
    extension = "pdf" if file_type == "pdf" else "xlsx"
    matches = list(OUTPUT_DIR.glob(f"{report_id}_*.{extension}"))
    return matches[0] if matches else None


@app.errorhandler(RequestEntityTooLarge)
def handle_large_upload(_error):
    flash(f"Upload is too large. Maximum total upload size is {MAX_UPLOAD_MB} MB.", "error")
    return redirect(url_for("index"))


@app.route("/", methods=["GET", "POST"])
def index():
    cleanup_old_files()
    download_links = None
    report_meta = None

    if request.method == "POST":
        cumulative_file = request.files.get("cumulative_csv")
        daily_file = request.files.get("daily_csv")

        try:
            if not cumulative_file or not daily_file:
                raise ValueError("Please upload both cumulative and daily CSV files.")

            cumulative_path = save_upload(cumulative_file, "cumulative")
            daily_path = save_upload(daily_file, "daily")

            result = generate_report(cumulative_path, daily_path, OUTPUT_DIR)
            report_id = uuid.uuid4().hex
            pdf_path = OUTPUT_DIR / f"{report_id}_{result['pdf_path'].name}"
            excel_path = OUTPUT_DIR / f"{report_id}_{result['excel_path'].name}"
            result["pdf_path"].rename(pdf_path)
            result["excel_path"].rename(excel_path)

            download_links = {
                "pdf": url_for("download_report", report_id=report_id, file_type="pdf"),
                "excel": url_for("download_report", report_id=report_id, file_type="excel"),
            }
            report_meta = {
                "generated": f"{result['title_date']} {result['title_time']}",
                "period": result["previous_period"],
            }
            flash("Report generated successfully.", "success")
        except (ValueError, ReportGenerationError) as exc:
            flash(str(exc), "error")
        except Exception:
            app.logger.exception("Report generation failed")
            flash("Report generation failed. Please check that both CSV files have the required NCRP columns.", "error")

    return render_template("index.html", download_links=download_links, report_meta=report_meta)


@app.route("/download/<report_id>/<file_type>")
def download_report(report_id, file_type):
    cleanup_old_files()
    safe_report_id = secure_filename(report_id)
    path = output_path_for(safe_report_id, file_type)
    if not path or not path.exists():
        flash("The requested report file has expired or does not exist. Please generate it again.", "error")
        return redirect(url_for("index"))

    download_name = "Daily_Report_GKR.pdf" if file_type == "pdf" else "Daily_Report_GKR.xlsx"
    return send_file(path, as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
