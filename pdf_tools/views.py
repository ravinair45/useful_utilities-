import os
import subprocess
from django.shortcuts import render, redirect
from django.http import HttpResponse, FileResponse
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from io import BytesIO
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from django.conf import settings
from django.core.files.storage import default_storage
import platform
import shutil
import fitz 
from zipfile import ZipFile
import uuid

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
        return FileResponse(open(file_path, 'rb'), as_attachment=True, filename="merged.pdf")
    else:
        return render(request, 'pdf_tools/result.html', {"error": "File not found."})