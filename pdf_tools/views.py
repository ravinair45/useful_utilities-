import os, uuid
from django.shortcuts import render, redirect
from django.http import HttpResponse, FileResponse
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from io import BytesIO
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from django.conf import settings
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
import platform
import shutil
import fitz 
from zipfile import ZipFile
from PIL import Image
import yt_dlp

def dashboard(request):
    return render(request, "pdf_tools/dashboard.html")

def find_libreoffice_executable():
    """
    Detects LibreOffice executable depending on OS.
    Returns full path or raises FileNotFoundError.
    """
    system = platform.system()
    
    # macOS
    if system == "Darwin":
        possible_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(possible_path):
            return possible_path
    
    # Linux
    if system == "Linux":
        exe = shutil.which("libreoffice") or shutil.which("soffice")
        if exe:
            return exe
    
    # Windows
    if system == "Windows":
        program_files = os.environ.get("PROGRAMFILES", "C:\\Program Files")
        possible_path = os.path.join(program_files, "LibreOffice", "program", "soffice.exe")
        if os.path.exists(possible_path):
            return possible_path
    
    raise FileNotFoundError("LibreOffice executable not found. Please install LibreOffice and ensure it's in PATH.")

def merge_pdf(request):
    if request.method == 'POST':
        files = request.FILES.getlist('pdfs')
        if not files:
            return render(request, 'pdf_tools/merge_pdf.html', {"error": "No files uploaded."})

        merger = PdfMerger()
        for f in files:
            merger.append(f)

        # unique file name
        file_id = str(uuid.uuid4()) + ".pdf"
        output_path = os.path.join(settings.MEDIA_ROOT, file_id)

        # make sure MEDIA_ROOT exists
        os.makedirs(settings.MEDIA_ROOT, exist_ok=True)

        with open(output_path, 'wb') as output_file:
            merger.write(output_file)

        merger.close()

        # redirect to result page with filename
        return redirect('result', file_id=file_id)

    return render(request, 'pdf_tools/merge_pdf.html')

def split_pdf(request):
    if request.method == "POST":
        pdf_file = request.FILES.get("pdf")
        pages_str = request.POST.get("page_range", "").strip()

        if not pdf_file or not pages_str:
            return render(request, "pdf_tools/split_pdf.html", {
                "error": "Please upload a PDF and specify pages (e.g., 1,3,5-7)."
            })

        try:
            reader = PdfReader(pdf_file)
            writer = PdfWriter()

            # Parse user input: "1,3,5-7"
            pages = []
            for part in pages_str.split(","):
                part = part.strip()
                if "-" in part:
                    start, end = part.split("-")
                    pages.extend(range(int(start) - 1, int(end)))  # 1-based to 0-based
                else:
                    pages.append(int(part) - 1)

            # Add valid pages to writer
            for page_num in pages:
                if 0 <= page_num < len(reader.pages):
                    writer.add_page(reader.pages[page_num])

            # Save the new PDF in MEDIA_ROOT
            file_id = str(uuid.uuid4()) + ".pdf"
            output_path = os.path.join(settings.MEDIA_ROOT, file_id)
            os.makedirs(settings.MEDIA_ROOT, exist_ok=True)

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            # Redirect to result page with operation type
            return redirect(f"/result/{file_id}?op=Split PDF")

        except Exception as e:
            return render(request, "pdf_tools/split_pdf.html", {
                "error": f"Error processing PDF: {e}"
            })

    return render(request, "pdf_tools/split_pdf.html")

def pdf_to_image(request):
    if request.method == "POST":
        pdf_file = request.FILES.get("pdf_file")
        if not pdf_file:
            return HttpResponse("Please upload a PDF file.")

        try:
            # Open PDF with PyMuPDF
            pdf_doc = fitz.open(stream=pdf_file.read(), filetype="pdf")

            # Create a zip in memory
            zip_buffer = BytesIO()
            with ZipFile(zip_buffer, "w") as zip_file:
                for page_num in range(pdf_doc.page_count):
                    page = pdf_doc[page_num]
                    pix = page.get_pixmap()  # render page to image
                    img_bytes = pix.tobytes("png")
                    zip_file.writestr(f"page_{page_num+1}.png", img_bytes)

            pdf_doc.close()
            zip_buffer.seek(0)

            # Return zip file
            response = HttpResponse(zip_buffer, content_type="application/zip")
            response["Content-Disposition"] = 'attachment; filename="pdf_images.zip"'
            return response

        except Exception as e:
            return HttpResponse(f"Error converting PDF to images: {e}")

    return render(request, "pdf_tools/pdf_to_image.html")

def result(request, file_id):
    file_url = settings.MEDIA_URL + file_id
    return render(request, 'pdf_tools/result.html', {
        "file_url": file_url,
        "file_id": file_id,
    })


def download_file(request, file_id):
    file_path = os.path.join(settings.MEDIA_ROOT, file_id)
    if os.path.exists(file_path):
        return FileResponse(
            open(file_path, 'rb'),
            as_attachment=True,
            filename=file_id  # preserve original filename (UUID + ext)
        )
    else:
        return render(request, 'pdf_tools/result.html', {"error": "File not found."})

def compress_pdf(request):
    if request.method == "POST":
        pdf_file = request.FILES.get("pdf_file")
        compression_level = request.POST.get("compression_level")  # 'basic' or 'strong'

        if not pdf_file:
            return HttpResponse("Please upload a PDF file.")

        try:
            pdf_doc = fitz.open(stream=pdf_file.read(), filetype="pdf")

            output_pdf = fitz.open()

            for page_num in range(pdf_doc.page_count):
                page = pdf_doc[page_num]

                # Scale factor changes based on compression level
                if compression_level == "basic":
                    zoom = 1.0  # keep same resolution
                    quality = 75
                else:  # strong compression
                    zoom = 0.5  # reduce resolution by half
                    quality = 40

                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))

                # Convert pixmap to Pillow Image
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # Save compressed image to memory
                img_bytes_io = BytesIO()
                img.save(img_bytes_io, format="JPEG", quality=quality, optimize=True)
                img_bytes = img_bytes_io.getvalue()

                # Insert compressed image back into PDF
                rect = fitz.Rect(0, 0, pix.width, pix.height)
                new_page = output_pdf.new_page(width=rect.width, height=rect.height)
                new_page.insert_image(rect, stream=img_bytes)

            # Save compressed PDF
            file_id = f"{uuid.uuid4()}_compressed.pdf"
            output_path = os.path.join("media", file_id)
            output_pdf.save(output_path, deflate=True)

            pdf_doc.close()
            output_pdf.close()

            return redirect("result", file_id=file_id)

        except Exception as e:
            return HttpResponse(f"Error compressing PDF: {e}")

    return render(request, "pdf_tools/compress_pdf.html")

def youtube_download(request):
    if request.method == "POST":
        url = request.POST.get("url")
        format_choice = request.POST.get("format", "mp4")  # default: mp4
        if not url:
            return HttpResponse("Please provide a valid YouTube URL.")

        # Ensure MEDIA_ROOT/youtube exists
        download_path = os.path.join(settings.MEDIA_ROOT, "youtube")
        os.makedirs(download_path, exist_ok=True)

        try:
            # Unique filename ID
            file_id = str(uuid.uuid4())

            # Set correct output template
            if format_choice == "mp3":
                outtmpl = os.path.join(download_path, f"{file_id}.%(ext)s")
                ydl_opts = {
                    "outtmpl": outtmpl,
                    "format": "bestaudio/best",
                    "postprocessors": [{
                        "key": "FFmpegExtractAudio",
                        "preferredcodec": "mp3",
                        "preferredquality": "192",
                    }],
                }
            else:  # mp4 video
                outtmpl = os.path.join(download_path, f"{file_id}.%(ext)s")
                ydl_opts = {
                    "outtmpl": outtmpl,
                    "format": "mp4/best",
                }

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=True)
                filename = ydl.prepare_filename(info)

            # Detect actual extension (yt-dlp decides)
            _, ext = os.path.splitext(filename)
            final_file_id = file_id + ext

            # Rename to consistent file_id.ext inside MEDIA_ROOT
            final_path = os.path.join(settings.MEDIA_ROOT, final_file_id)
            os.rename(filename, final_path)

            # Redirect to existing result view
            return redirect("result", file_id=final_file_id)

        except Exception as e:
            return HttpResponse(f"Error: {str(e)}")

    return render(request, "pdf_tools/youtube_download.html")