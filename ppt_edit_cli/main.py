import click
import os
import glob
import shutil
from pathlib import Path
from dotenv import load_dotenv
from .engine import PPTEngine
from .utils import pdf_to_png

@click.command()
@click.option('--input-dir', '-i', required=True, type=click.Path(exists=True), help='Directory containing PPT/PDF images.')
@click.option('--output', '-o', default='output.pptx', help='Output PPTX file path.')
@click.option('--api-key', '-k', default=None, help='Gemini API Key for semantic reconstruction.')
def main(input_dir, output, api_key):
    """
    PPT-Edit-CLI: Reconstruct editable PPT from images.
    """
    load_dotenv()
    key = api_key or os.getenv('GEMINI_API_KEY')
    
    if not key:
        click.echo("Warning: No Gemini API Key provided. Semantic reconstruction will be limited.")
    
    click.echo(f"Initializing engine... (This can take 10-20s for the first time)")
    engine = PPTEngine(api_key=key)
    
    # If input is a PDF, convert it to images first
    temp_dir = None
    if input_dir.lower().endswith('.pdf'):
        click.echo(f"Converting PDF to images...")
        temp_dir = os.path.join(os.path.dirname(input_dir), "temp_slides")
        image_files = pdf_to_png(input_dir, output_dir=temp_dir)
    else:
        # Supported image formats
        extensions = ['*.png', '*.jpg', '*.jpeg', '*.bmp']
        image_files = []
        for ext in extensions:
            image_files.extend(glob.glob(os.path.join(input_dir, ext)))
    
    if not image_files:
        click.echo(f"No slides found to process.")
        return

    image_files.sort() # Ensure correct slide order
    click.echo(f"Found {len(image_files)} slides. Processing...")
    
    processed_slides = []
    with click.progressbar(image_files, label='Processing Slides') as bar:
        for img_path in bar:
            try:
                slide_data = engine.process_image(img_path)
                processed_slides.append(slide_data)
            except Exception as e:
                click.echo(f"\nError processing {img_path}: {e}")

    click.echo(f"Generating PPT: {output}...")
    try:
        engine.create_ppt(processed_slides, output)
        click.echo("Success! Your editable PPT is ready.")
        
        # Cleanup
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            click.echo("Cleaned up temporary slide images.")
            
    except Exception as e:
        click.echo(f"Failed to generate PPT: {e}")

if __name__ == '__main__':
    main()
