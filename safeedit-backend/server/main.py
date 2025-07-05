from flask import Flask, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import zipfile
import io
from pdf2docx import Converter as PDF2DOCX
from docx2pdf import convert as DOCX2PDF
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from PIL import Image
import platform
if platform.system() == "Windows":
    import pythoncom
from pdf2image import convert_from_path
import pytesseract
import fitz  
import time
import pikepdf

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
MERGE_FOLDER = os.path.join(UPLOAD_FOLDER, 'merge')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(MERGE_FOLDER, exist_ok=True)

# 🧠 Thread-safe DOCX to PDF
def convert_docx_thread_safe(input_path, output_path):
    import platform
    if platform.system() == "Windows":
        import pythoncom
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
    else:
        raise NotImplementedError("DOCX to PDF conversion is only supported on Windows.")


@app.route('/')
def home():
    return "✅ Flask server is running for SafeEditPDF"

@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('file')
    action = request.form.get('type')

    print(f"=== Received {len(files)} file(s)")
    print("=== Conversion Type:", action)

    if not files or not action:
        return "❌ Missing file(s) or conversion type", 400

    try:
        # ✅ Handle multi-file image to PDF conversion
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
                    return f"❌ Error reading image: {fname} - {str(e)}", 400

            if not image_list:
                return "❌ No valid images to convert", 400

            image_list[0].save(
                output_path,
                save_all=True,
                append_images=image_list[1:],
                resolution=100.0,
                format="PDF"
            )

            for img in image_list:
                img.close()

            # Clean up uploaded image files
            for f in os.listdir(UPLOAD_FOLDER):
                if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.webp')):
                    os.remove(os.path.join(UPLOAD_FOLDER, f))

            return send_file(output_path, as_attachment=True)

        # 🛑 Only enforce single file for the rest
        if len(files) != 1:
            return "❌ This conversion only supports 1 file", 400


        file = files[0]
        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)

        if action == "pdf-to-docx":
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".pdf", ".docx"))
            converter = PDF2DOCX(input_path)
            converter.convert(output_path)
            converter.close()

        elif action == "docx-to-pdf":
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".docx", ".pdf"))
            convert_docx_thread_safe(input_path, output_path)

        elif action == "img-to-pdf":
            try:
                image_list = []

                for idx, f in enumerate(files):
                    if f and f.filename.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
                        print(f"➡️ Processing image {idx+1}: {f.filename}")
                        image = Image.open(f.stream).convert("RGB")
                        image_list.append(image)
                    else:
                        print(f"⚠️ Skipped: {f.filename}")

                if not image_list:
                    return "❌ No valid images to convert", 400

                output_path = os.path.join(OUTPUT_FOLDER, "merged_images.pdf")
                image_list[0].save(
                    output_path,
                    save_all=True,
                    append_images=image_list[1:],
                    resolution=100.0,
                    format="PDF"
                )

                # Clean up file handles
                for img in image_list:
                    img.close()

                return send_file(output_path, as_attachment=True, download_name="merged_images.pdf")

            except Exception as img_err:
                print("❌ Image to PDF conversion failed:", str(img_err))
                return f"❌ Image to PDF conversion failed: {str(img_err)}", 500



        elif action == "compress-pdf":
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".pdf", "_compressed.pdf"))

            quality = request.form.get("quality", "/screen")  # Default to /screen

            try:
                gs_path = r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe"
                command = [
                    gs_path,
                    "-sDEVICE=pdfwrite",
                    "-dCompatibilityLevel=1.4",
                    "-dPDFSETTINGS=/screen",  # Try /screen or /ebook
                    "-dDownsampleColorImages=true",
                    "-dColorImageResolution=72",
                    "-dDownsampleGrayImages=true",
                    "-dGrayImageResolution=72",
                    "-dDownsampleMonoImages=true",
                    "-dMonoImageResolution=72",
                    "-dNOPAUSE",
                    "-dQUIET",
                    "-dBATCH",
                    f"-sOutputFile={output_path}",
                    input_path,
                ]


                import subprocess
                subprocess.run(command, check=True)

                return send_file(output_path, as_attachment=True)

            except subprocess.CalledProcessError as e:
                return f"❌ Ghostscript compression failed: {str(e)}", 500




        elif action == "ocr-pdf":
            output_path = os.path.join(OUTPUT_FOLDER, filename.replace(".pdf", ".txt"))
            pages = convert_from_path(input_path)
            full_text = ""

            for page in pages:
                text = pytesseract.image_to_string(page)
                full_text += text + "\n\n"

            with open(output_path, "w", encoding="utf-8") as f:
                f.write(full_text)

        elif action == "split-pdf":
            reader = PdfReader(input_path)
            zip_buffer = io.BytesIO()
            zipf = zipfile.ZipFile(zip_buffer, "w")

            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                temp_output = f"page_{i + 1}.pdf"
                temp_path = os.path.join(OUTPUT_FOLDER, temp_output)
                with open(temp_path, "wb") as f:
                    writer.write(f)
                zipf.write(temp_path, arcname=temp_output)
                os.remove(temp_path)

            zipf.close()
            zip_buffer.seek(0)
            return send_file(zip_buffer, download_name="split_pages.zip", as_attachment=True)

        else:
            return "❌ Invalid conversion type", 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"❌ Error during conversion: {str(e)}", 500


# ✅ NEW: Encrypt PDF with password protection
@app.route("/encrypt", methods=["POST"])
def encrypt_pdf():
    if "file" not in request.files or "password" not in request.form:
        return "Missing file or password", 400

    file = request.files["file"]
    password = request.form["password"]

    try:
        reader = PdfReader(file)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(user_password=password, owner_password=password, use_128bit=True)

        output_stream = io.BytesIO()
        writer.write(output_stream)
        output_stream.seek(0)

        return send_file(output_stream, as_attachment=True, download_name="protected.pdf", mimetype="application/pdf")
    except Exception as e:
        return str(e), 500


if __name__ == '__main__':
    app.run(port=5001, debug=True)
