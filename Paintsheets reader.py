#Paintsheets

import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os

# Path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Mapping of part numbers to materials (example)
part_number_to_materials = {
    "12345": ["Material A", "Material B"],
    "67890": ["Material C", "Material D"],
    # Add more part numbers and corresponding materials here
}

def expand_part_numbers(extracted_text, part_number_to_materials):
    words = extracted_text.split()
    expanded_text = []

    for word in words:
        if word in part_number_to_materials:
            materials = part_number_to_materials[word]
            expanded_text.append(f"{word} ({', '.join(materials)})")
        else:
            expanded_text.append(word)
    
    return " ".join(expanded_text)

def pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path)
    text_data = []
    
    for i, image in enumerate(images):
        text = pytesseract.image_to_string(image)
        expanded_text = expand_part_numbers(text, part_number_to_materials)
        text_data.append(expanded_text)
    
    return text_data

def save_to_excel(text_data, output_path):
    df = pd.DataFrame({'Page': range(1, len(text_data) + 1), 'Text': text_data})
    df.to_excel(output_path, index=False)

def process_pdfs_in_folder(folder_path, output_excel_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    
    all_text_data = []
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        text_data = pdf_to_text(pdf_path)
        all_text_data.extend(text_data)
    
    save_to_excel(all_text_data, output_excel_path)

# Usage
folder_path = r'Path\To\Your\PDFs'
output_excel_path = r'Path\To\Save\Output.xlsx'

process_pdfs_in_folder(folder_path, output_excel_path)


