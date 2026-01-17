# app.py
from flask import Flask, render_template, request, send_file
from pathlib import Path
import os
import tempfile
from werkzeug.utils import secure_filename

from web import convert_pdfs_to_excel  # <-- IMPORTANT: import from converter.py


app = Flask(__name__, template_folder="templates")

# OPTIONAL SAFETY LIMITS (tune these)
# Total request size (all PDFs combined). If exceeded, Flask returns 413.
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH", 25 * 1024 * 1024))  # 25MB default

MAX_FILES_PER_UPLOAD = int(os.environ.get("MAX_FILES_PER_UPLOAD", "10"))

BASE_UPLOAD_DIR = Path("uploads")
BASE_OUTPUT_DIR = Path("outputs")
BASE_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("pdfs")
        if not files or files[0].filename == "":
            return "No PDF files selected!", 400

        if len(files) > MAX_FILES_PER_UPLOAD:
            return f"Too many files. Max allowed: {MAX_FILES_PER_UPLOAD}", 400

        # Use a per-request temp folder to avoid collisions
        with tempfile.TemporaryDirectory(prefix="cogs_") as tmpdir:
            tmpdir_path = Path(tmpdir)
            upload_dir = tmpdir_path / "uploads"
            output_dir = tmpdir_path / "outputs"
            upload_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)

            pdf_paths = []
            for f in files:
                original = f.filename or "file.pdf"
                filename = secure_filename(original) or "file.pdf"

                # Basic validation
                if not filename.lower().endswith(".pdf"):
                    return f"Invalid file type: {filename}. Only PDFs allowed.", 400

                path = upload_dir / filename
                f.save(path)
                pdf_paths.append(path)

            print(f"[DEBUG] Saved PDFs: {[p.name for p in pdf_paths]}")

            try:
                excel_file = convert_pdfs_to_excel(pdf_paths, output_dir)
            except Exception as e:
                print(f"[ERROR] PDF to Excel conversion failed: {e}")
                return "Failed to process PDF files", 500

            if excel_file is None or not excel_file.exists():
                print(f"[ERROR] Excel file not found: {excel_file}")
                return "Failed to process PDF files", 500

            print(f"[INFO] Sending Excel file: {excel_file}")

            return send_file(
                str(excel_file.resolve()),
                as_attachment=True,
                download_name=excel_file.name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    return render_template("index.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # debug=False in production; Render uses gunicorn anyway
    app.run(host="0.0.0.0", port=port, debug=True)
