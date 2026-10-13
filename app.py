"""
रजत पीडीएफ कन्वर्टर (Rajat PDF Converter Pro)
पूर्ण रूपेण ऑफ़लाइन, तीव्र और सुरक्षित दस्तावेज़ कनवर्टर एवं पीडीएफ टूलकिट
"""

import os
import sys
import uuid
import zipfile
import shutil
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template, send_from_directory
from werkzeug.utils import secure_filename

# Import conversion engine
from converter_engine import (
    word_to_pdf,
    pdf_to_word,
    unlock_pdf,
    lock_pdf,
    merge_pdfs,
    split_pdf,
    compress_pdf,
    rotate_pdf,
    add_watermark,
    pdf_to_images,
    images_to_pdf,
    extract_text_and_info
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024  # 150 MB max limit

# In-memory history of recent conversions
CONVERSION_HISTORY = []

def record_history(tool_name, original_name, output_filename, original_size, output_size=0):
    entry = {
        "id": str(uuid.uuid4())[:8],
        "tool": tool_name,
        "original_name": original_name,
        "filename": output_filename,
        "time": datetime.now().strftime("%I:%M %p"),
        "date": datetime.now().strftime("%d-%b-%Y"),
        "size": format_size(output_size or 0)
    }
    CONVERSION_HISTORY.insert(0, entry)
    if len(CONVERSION_HISTORY) > 30:
        CONVERSION_HISTORY.pop()

def format_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/history", methods=["GET"])
def get_history():
    return jsonify({"success": True, "history": CONVERSION_HISTORY})

@app.route("/api/open-folder", methods=["POST"])
def open_output_folder():
    """Open the output folder in Windows Explorer."""
    try:
        os.startfile(OUTPUT_FOLDER)
        return jsonify({"success": True, "message": "आउटपुट फ़ोल्डर खुल गया है। (Folder opened)"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@app.route("/api/download/<filename>")
def download_file(filename):
    """Download converted file."""
    safe_filename = secure_filename(filename)
    file_path = os.path.join(OUTPUT_FOLDER, safe_filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "फाइल नहीं मिली (File not found)"}), 404
    return send_file(file_path, as_attachment=True, download_name=safe_filename)

@app.route("/api/preview/<filename>")
def preview_file(filename):
    """Serve file inline for browser/modal preview."""
    safe_filename = secure_filename(filename)
    file_path = os.path.join(OUTPUT_FOLDER, safe_filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "फाइल नहीं मिली"}), 404
    return send_file(file_path)

# ==========================================
# 1. Word to PDF
# ==========================================
@app.route("/api/convert/word-to-pdf", methods=["POST"])
def api_word_to_pdf():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया Word (.docx, .doc) फाइल चुनें।"}), 400
        file = request.files["file"]
        if not file.filename:
            return jsonify({"success": False, "message": "कोई फाइल नहीं चुनी गई।"}), 400

        orig_filename = secure_filename(file.filename)
        base_name, ext = os.path.splitext(orig_filename)
        if ext.lower() not in [".docx", ".doc"]:
            return jsonify({"success": False, "message": "कृपया केवल Word (.docx या .doc) फाइल अपलोड करें।"}), 400

        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        out_filename = f"{base_name}_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)
        file.save(in_path)

        res = word_to_pdf(in_path, out_path)
        out_size = os.path.getsize(out_path)
        record_history("Word ➜ PDF", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message", "Word फाइल सफलतापूर्वक PDF में बदल गई!"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 2. PDF to Word
# ==========================================
@app.route("/api/convert/pdf-to-word", methods=["POST"])
def api_pdf_to_word():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]
        if not file.filename:
            return jsonify({"success": False, "message": "कोई फाइल नहीं चुनी गई।"}), 400

        orig_filename = secure_filename(file.filename)
        base_name, ext = os.path.splitext(orig_filename)
        if ext.lower() != ".pdf":
            return jsonify({"success": False, "message": "कृपया केवल .pdf फाइल अपलोड करें।"}), 400

        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        out_filename = f"{base_name}_{uid}.docx"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)
        file.save(in_path)

        res = pdf_to_word(in_path, out_path)
        out_size = os.path.getsize(out_path)
        record_history("PDF ➜ Word", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message", "PDF सफलतापूर्वक Word (.docx) में बदल गई!"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 3. Unlock PDF (Remove Password)
# ==========================================
@app.route("/api/pdf/unlock", methods=["POST"])
def api_pdf_unlock():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया लॉक्ड PDF फाइल चुनें।"}), 400
        file = request.files["file"]
        password = request.form.get("password", "").strip()

        orig_filename = secure_filename(file.filename)
        base_name, _ = os.path.splitext(orig_filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        out_filename = f"{base_name}_unlocked_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)
        file.save(in_path)

        res = unlock_pdf(in_path, out_path, password)
        out_size = os.path.getsize(out_path)
        record_history("PDF अनलॉक (Unlock)", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message", "PDF का लॉक सफलतापूर्वक हटा दिया गया है!"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}"
        })
    except ValueError as ve:
        return jsonify({"success": False, "message": str(ve)}), 400
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 4. Lock PDF (Encrypt with Password)
# ==========================================
@app.route("/api/pdf/lock", methods=["POST"])
def api_pdf_lock():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]
        password = request.form.get("password", "").strip()
        if not password:
            return jsonify({"success": False, "message": "कृपया नया पासवर्ड दर्ज करें!"}), 400

        orig_filename = secure_filename(file.filename)
        base_name, _ = os.path.splitext(orig_filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        out_filename = f"{base_name}_protected_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)
        file.save(in_path)

        res = lock_pdf(in_path, out_path, password)
        out_size = os.path.getsize(out_path)
        record_history("PDF लॉक (Protect)", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message", "PDF में पासवर्ड लॉक लगा दिया गया है!"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 5. Merge Multiple PDFs
# ==========================================
@app.route("/api/pdf/merge", methods=["POST"])
def api_pdf_merge():
    try:
        files = request.files.getlist("files")
        if not files or len(files) < 2:
            return jsonify({"success": False, "message": "मर्ज करने के लिए कम से कम 2 PDF फाइलें चुनें।"}), 400

        saved_paths = []
        uid = uuid.uuid4().hex[:6]
        for f in files:
            safe_n = secure_filename(f.filename)
            p = os.path.join(UPLOAD_FOLDER, f"{uid}_{safe_n}")
            f.save(p)
            saved_paths.append(p)

        out_filename = f"Merged_Document_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)

        res = merge_pdfs(saved_paths, out_path)
        out_size = os.path.getsize(out_path)
        record_history("PDF मर्ज (Merge)", f"{len(files)} फाइलें", out_filename, 0, out_size)

        return jsonify({
            "success": True,
            "message": res.get("message", "फाइलें सफलतापूर्वक एक साथ जोड़ दी गईं!"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 6. Split PDF / Extract Pages
# ==========================================
@app.route("/api/pdf/split", methods=["POST"])
def api_pdf_split():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]
        page_range = request.form.get("page_range", "").strip()

        orig_filename = secure_filename(file.filename)
        base_name, _ = os.path.splitext(orig_filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        file.save(in_path)

        split_subdir = os.path.join(OUTPUT_FOLDER, f"split_{uid}")
        res = split_pdf(in_path, split_subdir, page_range)

        # If multiple files generated, package them into a zip
        gen_files = res["files"]
        if len(gen_files) == 1:
            out_filename = f"{base_name}_extracted_{uid}.pdf"
            final_path = os.path.join(OUTPUT_FOLDER, out_filename)
            shutil.move(gen_files[0], final_path)
            shutil.rmtree(split_subdir, ignore_errors=True)
            out_size = os.path.getsize(final_path)
        else:
            out_filename = f"{base_name}_all_pages_{uid}.zip"
            final_path = os.path.join(OUTPUT_FOLDER, out_filename)
            with zipfile.ZipFile(final_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for gf in gen_files:
                    zipf.write(gf, os.path.basename(gf))
            shutil.rmtree(split_subdir, ignore_errors=True)
            out_size = os.path.getsize(final_path)

        record_history("PDF स्प्लिट (Split)", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message", "पेज सफलतापूर्वक अलग कर लिए गए!"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}" if out_filename.endswith(".pdf") else None
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 7. Compress PDF
# ==========================================
@app.route("/api/pdf/compress", methods=["POST"])
def api_pdf_compress():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]

        orig_filename = secure_filename(file.filename)
        base_name, _ = os.path.splitext(orig_filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        out_filename = f"{base_name}_compressed_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)
        file.save(in_path)

        res = compress_pdf(in_path, out_path)
        out_size = os.path.getsize(out_path)
        record_history("PDF कंप्रेस (Compress)", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message"),
            "original_size": format_size(res["original_size"]),
            "compressed_size": format_size(res["compressed_size"]),
            "saved_percent": res["saved_percent"],
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 8. Rotate PDF Pages
# ==========================================
@app.route("/api/pdf/rotate", methods=["POST"])
def api_pdf_rotate():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]
        degrees = int(request.form.get("degrees", 90))

        orig_filename = secure_filename(file.filename)
        base_name, _ = os.path.splitext(orig_filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        out_filename = f"{base_name}_rotated_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)
        file.save(in_path)

        res = rotate_pdf(in_path, out_path, degrees)
        out_size = os.path.getsize(out_path)
        record_history(f"पेज रोटेट ({degrees}°)", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 9. Add Watermark
# ==========================================
@app.route("/api/pdf/watermark", methods=["POST"])
def api_pdf_watermark():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]
        watermark_text = request.form.get("text", "CONFIDENTIAL").strip()
        opacity = float(request.form.get("opacity", 0.25))

        orig_filename = secure_filename(file.filename)
        base_name, _ = os.path.splitext(orig_filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        out_filename = f"{base_name}_watermarked_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)
        file.save(in_path)

        res = add_watermark(in_path, out_path, watermark_text, opacity)
        out_size = os.path.getsize(out_path)
        record_history("वाटरमार्क (Watermark)", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": res.get("message"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 10. PDF to Images (PNG)
# ==========================================
@app.route("/api/convert/pdf-to-images", methods=["POST"])
def api_pdf_to_images():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]

        orig_filename = secure_filename(file.filename)
        base_name, _ = os.path.splitext(orig_filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        file.save(in_path)

        images_subdir = os.path.join(OUTPUT_FOLDER, f"img_{uid}")
        res = pdf_to_images(in_path, images_subdir)
        gen_imgs = res["images"]

        out_filename = f"{base_name}_images_{uid}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, out_filename)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for img_f in gen_imgs:
                zipf.write(img_f, os.path.basename(img_f))
        shutil.rmtree(images_subdir, ignore_errors=True)

        out_size = os.path.getsize(zip_path)
        record_history("PDF ➜ इमेज (Images)", orig_filename, out_filename, os.path.getsize(in_path), out_size)

        return jsonify({
            "success": True,
            "message": f"कुल {len(gen_imgs)} पेज इमेज (PNG) में बदलकर ZIP फाइल में तैयार कर दिए गए हैं!",
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 11. Images to PDF
# ==========================================
@app.route("/api/convert/images-to-pdf", methods=["POST"])
def api_images_to_pdf():
    try:
        files = request.files.getlist("files")
        if not files:
            return jsonify({"success": False, "message": "कृपया कम से कम एक इमेज फाइल (JPG/PNG) चुनें।"}), 400

        saved_images = []
        uid = uuid.uuid4().hex[:6]
        for f in files:
            safe_n = secure_filename(f.filename)
            p = os.path.join(UPLOAD_FOLDER, f"{uid}_{safe_n}")
            f.save(p)
            saved_images.append(p)

        out_filename = f"Images_Doc_{uid}.pdf"
        out_path = os.path.join(OUTPUT_FOLDER, out_filename)

        res = images_to_pdf(saved_images, out_path)
        out_size = os.path.getsize(out_path)
        record_history("इमेज ➜ PDF", f"{len(files)} फोटो", out_filename, 0, out_size)

        return jsonify({
            "success": True,
            "message": res.get("message"),
            "filename": out_filename,
            "size": format_size(out_size),
            "download_url": f"/api/download/{out_filename}",
            "preview_url": f"/api/preview/{out_filename}"
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

# ==========================================
# 12. Extract Text & Info
# ==========================================
@app.route("/api/pdf/extract-text", methods=["POST"])
def api_extract_text():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "message": "कृपया PDF फाइल चुनें।"}), 400
        file = request.files["file"]

        orig_filename = secure_filename(file.filename)
        uid = uuid.uuid4().hex[:6]
        in_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{orig_filename}")
        file.save(in_path)

        res = extract_text_and_info(in_path)
        return jsonify({
            "success": True,
            "page_count": res["page_count"],
            "full_text": res["full_text"],
            "metadata": res["metadata"]
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"त्रुटि: {str(e)}"}), 500

if __name__ == "__main__":
    print("==================================================")
    print(" 🚀 रजत पीडीएफ कन्वर्टर (Rajat PDF Converter Pro)")
    print(" सर्वर चालू हो गया है: http://127.0.0.1:5050")
    print("==================================================")
    app.run(host="127.0.0.1", port=5050, debug=False)
