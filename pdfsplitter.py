import fitz  # PyMuPDF
import os
from PIL import Image

# --------------------------
# SETTINGS
pdf_file = "BPMN ESEM Elnätsanslutningar.drawio-1-1_rotated.pdf"
output_folder = r"C:\Users\Mohammed\Desktop\pdfconverter\output_pages"

dpi = 200                       # Skarpt men rimlig filstorlek

# Word / Google Docs – Letter
page_width_in = 8.5
page_height_in = 11
# --------------------------

# Inches → points
page_width_pts = page_width_in * 72
page_height_pts = page_height_in * 72

os.makedirs(output_folder, exist_ok=True)

doc = fitz.open(pdf_file)
page = doc[0]

orig_width = page.rect.width
orig_height = page.rect.height

# Skala så att bredden matchar Word-sida exakt
scale = page_width_pts / orig_width

# Hur mycket av PDF:en som ryms på en Word-sida
slice_height_pdf = page_height_pts / scale

num_slices = int(orig_height // slice_height_pdf) + 1
print(f"Slicing into {num_slices} Word-sized pages")

matrix = fitz.Matrix(dpi / 72 * scale, dpi / 72 * scale)

for i in range(num_slices):
    top = i * slice_height_pdf
    bottom = min((i + 1) * slice_height_pdf, orig_height)

    clip = fitz.Rect(0, top, orig_width, bottom)
    pix = page.get_pixmap(matrix=matrix, clip=clip)

    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    output_file = os.path.join(output_folder, f"page_{i+1}.png")
    img.save(output_file, dpi=(dpi, dpi))
    print(f"Saved {output_file}")

print("✅ Klar! PDFen är skuren i Word-/Google Docs-sidor.")
