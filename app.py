# app.py
from flask import Flask, render_template, request, send_file
from pathlib import Path
import os
from web import convert_pdfs_to_excel  # your converter.py

# ---------------------------
# App setup
# ---------------------------
app = Flask(__name__, template_folder="templates")

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------
# Routes
# ---------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("pdfs")
        if not files or files[0].filename == "":
            return "No PDF files selected!", 400

        pdf_paths = []
        for f in files:
            path = UPLOAD_DIR / f.filename
            f.save(path)
            pdf_paths.append(path)
        print(f"[DEBUG] Saved PDFs: {[p.name for p in pdf_paths]}")

        # Convert PDFs to Excel
        try:
            excel_file = convert_pdfs_to_excel(pdf_paths, OUTPUT_DIR)
        except Exception as e:
            print(f"[ERROR] PDF to Excel conversion failed: {e}")
            return "Failed to process PDF files", 500

        # Clean uploaded PDFs after processing
        for pdf_path in pdf_paths:
            pdf_path.unlink(missing_ok=True)
        print(f"[DEBUG] Cleaned uploaded PDFs")

        if excel_file is None or not excel_file.exists():
            print(f"[ERROR] Excel file not found: {excel_file}")
            return "Failed to process PDF files", 500

        print(f"[INFO] Sending Excel file: {excel_file}")
        return send_file(
            str(excel_file.resolve()),  # absolute path
            as_attachment=True,
            download_name=excel_file.name
        )

    return render_template("index.html")


# ---------------------------
# Run (local testing)
# ---------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
