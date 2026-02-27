import fitz  # PyMuPDF
import os
from pathlib import Path
from pptx.util import Inches, Emu

def px_to_emu(px, total_px, slide_length_inches):
    """
    Convert pixels to EMU based on the total pixels and slide length in inches.
    1 Inch = 914,400 EMUs.
    """
    inches = (px / total_px) * slide_length_inches
    return Emu(int(inches * 914400))

def hex_to_rgb(hex_color):
    """Convert hex color string to RGB tuple."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def pdf_to_png(pdf_path, output_dir=None, dpi=200):
    """
    Convert PDF pages to PNG images.
    """
    pdf_doc = fitz.open(pdf_path)
    if output_dir is None:
        output_dir = Path(pdf_path).parent / "temp_slides"
    else:
        output_dir = Path(output_dir)
        
    output_dir.mkdir(parents=True, exist_ok=True)
    
    png_files = []
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    
    for i in range(len(pdf_doc)):
        page = pdf_doc.load_page(i)
        pix = page.get_pixmap(matrix=mat)
        output_path = output_dir / f"page_{i+1:03d}.png"
        pix.save(str(output_path))
        png_files.append(str(output_path))
        
    pdf_doc.close()
    return png_files
