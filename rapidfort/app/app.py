import os
import pythoncom
from flask import Flask, request, render_template, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter

app = Flask(__name__)
app.secret_key = "your_secret_key"
UPLOAD_FOLDER = "uploads"
PDF_FOLDER = "pdfs"

# Create necessary folders
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["PDF_FOLDER"] = PDF_FOLDER

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        flash("No file part")
        return redirect(url_for("index"))

    file = request.files["file"]
    if file.filename == "":
        flash("No file selected")
        return redirect(url_for("index"))

    if file and file.filename.endswith(".docx"):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(file_path)

        # Initialize COM before calling docx2pdf
        pythoncom.CoInitialize()

        try:
            # Convert DOCX to PDF
            pdf_filename = filename.replace(".docx", ".pdf")
            pdf_path = os.path.join(app.config["PDF_FOLDER"], pdf_filename)
            convert(file_path, pdf_path)

        finally:
            # Uninitialize COM after the operation
            pythoncom.CoUninitialize()

        # Optionally gather metadata here, if needed (e.g., file size)
        metadata = {
            "Filename": pdf_filename,
            "File Size": os.path.getsize(pdf_path),
        }

        # Render the result page, passing the PDF path and metadata
        return render_template("result.html", pdf_path=pdf_path, metadata=metadata)

    flash("Only .docx files are allowed!")
    return redirect(url_for("index"))

@app.route("/result", methods=["POST"])
def result():
    pdf_path = request.form.get("pdf_path")
    password_option = request.form.get("password_option")
    
    if password_option == "protected":
        password = request.form.get("password")
        
        # Add password protection to PDF using PyPDF2
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        # Add all pages to writer
        for page_num in range(len(reader.pages)):
            writer.add_page(reader.pages[page_num])

        # Encrypt the PDF with the chosen password
        protected_pdf_path = os.path.join(app.config["PDF_FOLDER"], f"protected_{os.path.basename(pdf_path)}")
        with open(protected_pdf_path, "wb") as f:
            writer.encrypt(password)
            writer.write(f)

        return send_file(protected_pdf_path, as_attachment=True)

    # If no password protection is selected, simply return the normal PDF
    return send_file(pdf_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
