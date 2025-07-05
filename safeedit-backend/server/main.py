from flask import Flask, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import io
import zipfile
import time
import subprocess
import platform

from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter as PDF2DOCX
from docx2pdf import convert as DOCX2PDF
from PIL import Image
from pdf2image import convert_from_path
import pytesseract

if platform.system() == "Windows":
    import pythoncom

# ✅ Folder setup
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
MERGE_FOLDER = os.path.join(UPLOAD_FOLDER, "merge")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(MERGE_FOLDER, exist_ok=True)

app = Flask(__name__)
CORS(app)



# 🧠 Safe DOCX ➡ PDF (Windows only)
def convert_docx_thread_safe(input_path, output_path):
    system = platform.system()

    if system == "Windows":
        import pythoncom
        from docx2pdf import convert as DOCX2PDF

        for attempt in range(3):
            pythoncom.CoInitialize()
            try:
                DOCX2PDF(input_path, output_path)
                return
            except Exception as e:
                print(f"⚠️ Attempt {attempt+1} failed: {e}")
                time.sleep(1)
            finally:
                pythoncom.CoUninitialize()

        raise Exception("❌ DOCX to PDF conversion failed after 3 attempts.")

    elif system == "Linux":
        # Use LibreOffice CLI for conversion
        try:
            result = subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(output_path), input_path],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            print(result.stdout.decode())
        except subprocess.CalledProcessError as e:
            print("❌ LibreOffice failed:", e.stderr.decode())
            raise Exception(f"❌ LibreOffice conversion failed: {e.stderr.decode()}")
    else:
        raise NotImplementedError(f"DOCX to PDF is not supported on this OS: {system}")
        

@app.route('/')
def home():
    return "✅ SafeEditPDF Flask Server is running"

@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('file')
    action = request.form.get('type')

    if not files or not action:
        return "❌ Missing file(s) or conversion type", 400

    try:
        # ✅ Image(s) to PDF
        if action == "img-to-pdf":
            output_path = os.path.join(OUTPUT_FOLDER, "merged_images.pdf")
            image_list = []

            for f in files:
                fname = secure_filename(f.filename)
                path = os.path.join(UPLOAD_FOLDER, fname)
                f.save(path)
                try:
                    image = Image.open(path).convert("RGB")
                    image_list.append(image)
                except Exception as e:
                    return f"❌ Error reading image {fname}: {str(e)}", 400

            if not image_list:
                return "❌ No valid images found", 400

            image_list[0].save(
                output_path,
                save_all=True,
                append_images=image_list[1:],
                resolution=100.0,
                format="PDF"
            )

            for img in image_list:
                img.close()

            return send_file(output_path, as_attachment=True, download_name="merged_images.pdf")

        # 🛑 All others expect one file
        if len(files) != 1:
            return "❌ This conversion only supports 1 file", 400

        file = files[0]
        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)

        # ✅ PDF ➡ DOCX
        if action == "pdf-to-docx":
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".pdf", ".docx"))
            converter = PDF2DOCX(input_path)
            converter.convert(output_path)
            converter.close()

        # ✅ DOCX ➡ PDF (Windows only)
        elif action == "docx-to-pdf":
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".docx", ".pdf"))
            convert_docx_thread_safe(input_path, output_path)

        # ✅ Compress PDF
        elif action == "compress-pdf":
            # Get compression quality from form (default to /screen)
            quality = request.form.get("quality", "/screen")
        
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".pdf", "_compressed.pdf"))
            gs_cmd = "gswin64c" if platform.system() == "Windows" else "gs"
            command = [
                gs_cmd,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={quality}",  # ✅ Dynamic quality
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={output_path}",
                input_path
            ]
            try:
                subprocess.run(command, check=True)
                return send_file(output_path, as_attachment=True)
            except Exception as e:
                return f"❌ Ghostscript compression failed: {str(e)}", 500


        # ✅ OCR PDF ➡ Text
        elif action == "ocr-pdf":
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".pdf", ".txt"))
            pages = convert_from_path(input_path)
            full_text = ""
            for page in pages:
                text = pytesseract.image_to_string(page)
                full_text += text + "\n\n"
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(full_text)
        
            return send_file(output_path, as_attachment=True)  # ✅ Send .txt file as download


        # ✅ Split PDF ➡ ZIP
        elif action == "split-pdf":
            reader = PdfReader(input_path)
            zip_buffer = io.BytesIO()
            zipf = zipfile.ZipFile(zip_buffer, "w")
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                split_path = os.path.join(OUTPUT_FOLDER, f"page_{i+1}.pdf")
                with open(split_path, "wb") as f:
                    writer.write(f)
                zipf.write(split_path, arcname=os.path.basename(split_path))
                os.remove(split_path)
            zipf.close()
            zip_buffer.seek(0)
            return send_file(zip_buffer, download_name="split_pages.zip", as_attachment=True)

        else:
            return "❌ Invalid conversion type", 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"❌ Error during conversion: {str(e)}", 500

@app.route("/encrypt", methods=["POST"])
def encrypt_pdf():
    if "file" not in request.files or "password" not in request.form:
        return "❌ Missing file or password", 400

    file = request.files["file"]
    password = request.form["password"]

    try:
        reader = PdfReader(file)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(user_password=password, owner_password=password, use_128bit=True)

        stream = io.BytesIO()
        writer.write(stream)
        stream.seek(0)

        return send_file(stream, as_attachment=True, download_name="protected.pdf", mimetype="application/pdf")
    except Exception as e:
        return f"❌ Encryption error: {str(e)}", 500

if __name__ == "__main__":
    app.run(port=5001, debug=True)
