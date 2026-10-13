"""
रजत पीडीएफ कन्वर्टर (Rajat PDF Converter) - Core Conversion & PDF Engine
Supports:
1. Word to PDF (.docx, .doc -> .pdf)
2. PDF to Word (.pdf -> .docx)
3. Unlock PDF (पासवर्ड हटाएं / रिमूव पासवर्ड)
4. Lock PDF (पासवर्ड लगाएं / प्रोटेक्ट करें)
5. Merge PDFs (कई पीडीएफ जोड़ें)
6. Split & Extract Pages (पेज अलग करें)
7. Compress PDF (साइज कम करें)
8. PDF to Images (PNG/JPG में बदलें)
9. Images to PDF (फोटो से पीडीएफ बनाएं)
10. Rotate PDF (पेज घुमाएं)
11. Watermark PDF (वाटरमार्क लगाएं)
12. Extract Text & Metadata (टेक्स्ट निकालें)
"""

import os
import io
import math
import shutil
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def word_to_pdf(input_path, output_path):
    """
    Convert Word (.docx, .doc) to PDF.
    Primary: Microsoft Word COM Automation (Pixel-Perfect fidelity).
    Fallback: python-docx + ReportLab layout engine.
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    
    # Try Windows MS Word COM
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(input_path, ReadOnly=True)
            # 17 is wdFormatPDF in Word COM
            doc.SaveAs2(output_path, FileFormat=17)
            doc.Close(False)
            word.Quit()
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                return {"success": True, "message": "Microsoft Word इंजन द्वारा शत-प्रतिशत शुद्धता के साथ PDF में परिवर्तित किया गया।"}
        except Exception as e:
            try:
                word.Quit()
            except Exception:
                pass
            raise e
    except Exception as com_err:
        # Fallback using python-docx and reportlab if docx
        try:
            import docx
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            
            doc = docx.Document(input_path)
            styles = getSampleStyleSheet()
            story = []
            for para in doc.paragraphs:
                if para.text.strip():
                    story.append(Paragraph(para.text, styles['Normal']))
                    story.append(Spacer(1, 6))
            if not story:
                story.append(Paragraph("Empty Document", styles['Normal']))
            pdf_doc = SimpleDocTemplate(output_path, pagesize=letter)
            pdf_doc.build(story)
            return {"success": True, "message": "सफलतापूर्वक PDF में परिवर्तित किया गया (लेआउट इंजन द्वारा)।"}
        except Exception as fallback_err:
            raise RuntimeError(f"Word to PDF रूपांतरण में त्रुटि: {str(com_err)} | फॉलबैक: {str(fallback_err)}")

def pdf_to_word(input_path, output_path, start_page=0, end_page=None):
    """
    Convert PDF to Word (.docx) with tables, fonts, formatting, images and text.
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    try:
        cv = Converter(input_path)
        cv.convert(output_path, start=start_page, end=end_page)
        cv.close()
        return {"success": True, "message": "PDF फाइल सफलतापूर्वक Word (.docx) में बदल दी गई है!"}
    except Exception as e:
        raise RuntimeError(f"PDF से Word रूपांतरण में त्रुटि: {str(e)}")

def unlock_pdf(input_path, output_path, password=""):
    """
    Remove password from encrypted PDF and save an unlocked copy.
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    password = password or ""
    
    # Method 1: Using PyMuPDF (fitz)
    try:
        doc = fitz.open(input_path)
        if doc.is_encrypted:
            authenticated = doc.authenticate(password)
            if not authenticated:
                doc.close()
                raise ValueError("पासवर्ड गलत है! कृपया सही पासवर्ड डालें। (Incorrect Password)")
        doc.save(output_path, garbage=3, deflate=True)
        doc.close()
        return {"success": True, "message": "PDF का लॉक सफलतापूर्वक हटा दिया गया है! अब यह बिना पासवर्ड खुलेगी।"}
    except ValueError:
        raise
    except Exception:
        # Method 2: Fallback using pypdf
        try:
            reader = PdfReader(input_path)
            if reader.is_encrypted:
                res = reader.decrypt(password)
                if res == 0:
                    raise ValueError("पासवर्ड गलत है! कृपया सही पासवर्ड डालें। (Incorrect Password)")
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            with open(output_path, "wb") as f:
                writer.write(f)
            return {"success": True, "message": "PDF का लॉक सफलतापूर्वक हटा दिया गया है!"}
        except Exception as e:
            raise RuntimeError(f"PDF अनलॉक करने में त्रुटि: {str(e)}")

def lock_pdf(input_path, output_path, user_password, owner_password=None):
    """
    Encrypt PDF with strong password protection.
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    if not user_password:
        raise ValueError("कृपया नया पासवर्ड दर्ज करें! (Password is required)")
    
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    
    owner_pwd = owner_password if owner_password else user_password
    writer.encrypt(user_password=user_password, owner_password=owner_pwd, algorithm="AES-256")
    
    with open(output_path, "wb") as f:
        writer.write(f)
    return {"success": True, "message": "PDF में पासवर्ड लॉक सफलतापूर्वक लगा दिया गया है!"}

def merge_pdfs(file_paths, output_path):
    """
    Merge multiple PDF files into one combined PDF.
    """
    try:
        writer = PdfWriter()
        for fp in file_paths:
            writer.append(fp)
        with open(output_path, "wb") as f:
            writer.write(f)
        return {"success": True, "message": f"{len(file_paths)} PDF फाइलें सफलतापूर्वक एक साथ जोड़ दी गईं!"}
    except Exception as e:
        # Fallback with PyMuPDF
        try:
            merged_doc = fitz.open()
            for fp in file_paths:
                sub_doc = fitz.open(fp)
                merged_doc.insert_pdf(sub_doc)
                sub_doc.close()
            merged_doc.save(output_path)
            merged_doc.close()
            return {"success": True, "message": f"{len(file_paths)} PDF फाइलें सफलतापूर्वक एक साथ जोड़ दी गईं!"}
        except Exception as e2:
            raise RuntimeError(f"PDF जोड़ने में त्रुटि: {str(e)} | {str(e2)}")

def split_pdf(input_path, output_dir, page_ranges=""):
    """
    Split PDF or extract specific page ranges (e.g., '1-3, 5, 8').
    """
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    os.makedirs(output_dir, exist_ok=True)
    generated_files = []
    
    if page_ranges and page_ranges.strip():
        selected_indices = set()
        parts = [p.strip() for p in page_ranges.split(",") if p.strip()]
        for part in parts:
            if "-" in part:
                start, end = part.split("-", 1)
                try:
                    s = max(1, int(start.strip()))
                    e = min(total_pages, int(end.strip()))
                    for p in range(s, e + 1):
                        selected_indices.add(p - 1)
                except ValueError:
                    continue
            else:
                try:
                    p = int(part)
                    if 1 <= p <= total_pages:
                        selected_indices.add(p - 1)
                except ValueError:
                    continue
        
        sorted_indices = sorted(list(selected_indices))
        if not sorted_indices:
            raise ValueError("अमान्य पेज संख्या! कृपया मान्य पेज नंबर या रेंज दर्ज करें (उदा. 1-3, 5)")
        
        writer = PdfWriter()
        for idx in sorted_indices:
            writer.add_page(reader.pages[idx])
        out_file = os.path.join(output_dir, "extracted_pages.pdf")
        with open(out_file, "wb") as f:
            writer.write(f)
        generated_files.append(out_file)
        return {"success": True, "files": generated_files, "message": f"{len(sorted_indices)} पेज सफलतापूर्वक अलग (एक्सट्रैक्ट) कर लिए गए!"}
    else:
        # Split all pages individually
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        for idx, page in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(page)
            page_filename = os.path.join(output_dir, f"{base_name}_page_{idx + 1}.pdf")
            with open(page_filename, "wb") as f:
                writer.write(f)
            generated_files.append(page_filename)
        return {"success": True, "files": generated_files, "message": f"कुल {total_pages} पेज अलग-अलग फाइलों में विभाजित कर दिए गए!"}

def compress_pdf(input_path, output_path, quality="medium"):
    """
    Compress and reduce PDF file size while keeping text and graphics crisp.
    """
    orig_size = os.path.getsize(input_path)
    doc = fitz.open(input_path)
    doc.save(
        output_path,
        garbage=4,
        deflate=True,
        clean=True,
        deflate_images=True,
        deflate_fonts=True
    )
    doc.close()
    
    new_size = os.path.getsize(output_path)
    saved_bytes = max(0, orig_size - new_size)
    percent = (saved_bytes / orig_size * 100) if orig_size > 0 else 0
    
    return {
        "success": True,
        "original_size": orig_size,
        "compressed_size": new_size,
        "saved_percent": round(percent, 1),
        "message": f"PDF सफलतापूर्वक कंप्रेस हो गई! साइज {round(percent, 1)}% कम हो गया।"
    }

def rotate_pdf(input_path, output_path, degrees=90):
    """
    Rotate PDF pages (degrees: 90, 180, 270).
    """
    doc = fitz.open(input_path)
    for page in doc:
        page.set_rotation((page.rotation + degrees) % 360)
    doc.save(output_path)
    doc.close()
    return {"success": True, "message": f"सभी पेज {degrees}° घुमा दिए गए हैं!"}

def add_watermark(input_path, output_path, watermark_text, opacity=0.25):
    """
    Add stylish semi-transparent diagonal text watermark to each page.
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()
    
    for page in reader.pages:
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(page_width, page_height))
        can.setFont("Helvetica-Bold", 44)
        can.setFillColorRGB(0.4, 0.4, 0.4, alpha=opacity)
        
        can.saveState()
        can.translate(page_width / 2.0, page_height / 2.0)
        can.rotate(45)
        can.drawCentredString(0, 0, watermark_text)
        can.restoreState()
        can.save()
        
        packet.seek(0)
        watermark_pdf_obj = PdfReader(packet)
        watermark_page = watermark_pdf_obj.pages[0]
        
        page.merge_page(watermark_page)
        writer.add_page(page)
        
    with open(output_path, "wb") as f:
        writer.write(f)
    return {"success": True, "message": f"वाटरमार्क '{watermark_text}' सफलतापूर्वक लगाया गया!"}

def pdf_to_images(input_path, output_dir, zoom=2.0):
    """
    Render PDF pages into high-resolution PNG images.
    """
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(input_path)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    generated_images = []
    
    mat = fitz.Matrix(zoom, zoom)
    for page_idx in range(len(doc)):
        page = doc.load_page(page_idx)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_filename = f"{base_name}_page_{page_idx + 1}.png"
        img_filepath = os.path.join(output_dir, img_filename)
        pix.save(img_filepath)
        generated_images.append(img_filepath)
    doc.close()
    return {"success": True, "images": generated_images, "message": f"{len(generated_images)} पेज हाई-क्वालिटी PNG इमेज में कन्वर्ट हो गए!"}

def images_to_pdf(image_paths, output_path):
    """
    Convert image files (JPG, PNG, WEBP) into a single PDF.
    """
    if not image_paths:
        raise ValueError("कोई इमेज फाइल प्राप्त नहीं हुई।")
    
    images = []
    for path in image_paths:
        img = Image.open(path)
        if img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)
        
    images[0].save(output_path, save_all=True, append_images=images[1:], quality=95)
    return {"success": True, "message": f"{len(images)} फोटो से सफलतापूर्वक PDF बन गई!"}

def extract_text_and_info(input_path):
    """
    Extract text and metadata from PDF.
    """
    doc = fitz.open(input_path)
    total_pages = len(doc)
    text_content = []
    for i, page in enumerate(doc):
        text_content.append(f"--- [पेज {i + 1}] ---\n" + page.get_text())
    
    meta = doc.metadata or {}
    doc.close()
    return {
        "success": True,
        "page_count": total_pages,
        "metadata": meta,
        "full_text": "\n".join(text_content)
    }
