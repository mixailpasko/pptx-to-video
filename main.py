import os
import requests
from pptx import Presentation
from comtypes.client import CreateObject
from PIL import Image
import fitz # PyMuPD
from pytesseract import image_to_string
from moviepy import ImageClip, concatenate_videoclips, AudioFileClip
import inspect


# Step 1: Convert PPT to PDF
def ppt_to_pdf(ppt_path, output_pdf_path):
    print("Converting PPT to PDF...")
    powerpoint = CreateObject("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
    presentation.SaveAs(output_pdf_path, 32)
    presentation.Close()
    powerpoint.Quit()
    print(f"PPT converted to PDF: {output_pdf_path}")

def pdf_to_images_and_text(pdf_path, images_output_folder, text_output_folder):
    print("Splitting PDF into pages and extracting text...")
    os.makedirs(images_output_folder, exist_ok=True)
    os.makedirs(text_output_folder, exist_ok=True)
    pdf_document = fitz.open(pdf_path)
    texts = []

    for page_num in range(len(pdf_document)):
        # Save page as image
        page = pdf_document[page_num]
        pix = page.get_pixmap()
        image_path = os.path.join(images_output_folder, f"page_{page_num + 1}.png")
        pix.save(image_path)

        # Extract text using pytesseract
        image = Image.open(image_path)
        text = image_to_string(image, lang="eng")
        text_path = os.path.join(text_output_folder, f"page_{page_num + 1}.txt")
        with open(text_path, "w", encoding="utf-8") as f:
            f.write(text)
        texts.append((image_path, text))

    pdf_document.close()
    print("PDF processing completed.")
    return texts

def main():
    # Get the directory where the script is located and Construct the full path to the PPT file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ppt_path = os.path.join(script_dir, "test.pptx")
    # Construct the full path to the pdf file
    pdf_path = os.path.join(script_dir, "output.pdf")
    images_output_folder = os.path.join(script_dir, "path_to_images")
    text_output_folder = os.path.join(script_dir, "path_to_texts")

    ppt_to_pdf(ppt_path, pdf_path)

    help(image_to_string)

    print(inspect.getsource(image_to_string))

    

    # pdf_to_images_and_text(pdf_path, images_output_folder, text_output_folder)

if __name__ == "__main__":
    main()

