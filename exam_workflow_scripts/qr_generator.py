import fitz # PyMuPDF
import qrcode
import os
import io
from PIL import Image
import shutil
import json
import sys
# Import from shared utility - now csv_utils
from exam_workflow_scripts.csv_utils import load_students_from_csv_files

# --- CONFIGURATION ---
FITXER_BASE = "base.pdf" 
TEXT_A_BUSCAR = "%NOM%"
OUTPUT_EXAMS_DIR = "generated_exams" 
OUTPUT_ZIP_DIR = "zipped_exams"
STUDENT_DATA_DIR = "student_data" # New directory for student CSVs

MIDA_QR = 25
POS_QR_X = 20
POS_QR_Y = 800

FONT_TEXT = "helv"
MIDA_TEXT = 12

def generar_imatge_qr(dades):
    """Generates a QR code and returns it as bytes."""
    qr = qrcode.QRCode(box_size=10, border=0)
    qr.add_data(dades)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

def generate_individual_exams(base_pdf_path, students_data, output_dir_for_exams): # Renamed parameter for clarity
    """
    Generates individual PDF exams for each student with their name and a unique QR code.
    """
    os.makedirs(output_dir_for_exams, exist_ok=True) # Use renamed parameter
    print(f"Processing {len(students_data)} exams...")

    if not os.path.exists(base_pdf_path):
        raise FileNotFoundError(f"Base PDF file not found: '{base_pdf_path}'. Please ensure it exists.")

    try:
        doc_test = fitz.open(base_pdf_path)
        doc_test.close()
    except Exception as e:
        raise Exception(f"Error opening base PDF file '{base_pdf_path}': {e}")

    for alumne in students_data:
        doc = fitz.open(base_pdf_path) # Open a fresh copy for each student

        for page_num, page in enumerate(doc):
            # Find and replace student name placeholder
            areas_trobades = page.search_for(TEXT_A_BUSCAR)
            for rect in areas_trobades:
                adjusted_rect = fitz.Rect(rect.x0, rect.y0 + 3, rect.x1, rect.y1 - 3)
                page.draw_rect(adjusted_rect, color=(1, 1, 1), fill=(1, 1, 1))
                page.insert_text((rect.x0, rect.y1 - 2), alumne['nom'], fontsize=MIDA_TEXT, fontname=FONT_TEXT, color=(0, 0, 0))

            # Insert QR code
            qr_content = f'{alumne["id"]}-{page_num + 1}'
            qr_bytes = generar_imatge_qr(qr_content)
            rect_qr = fitz.Rect(POS_QR_X, POS_QR_Y, POS_QR_X + MIDA_QR, POS_QR_Y + MIDA_QR)
            page.insert_image(rect_qr, stream=qr_bytes)

        # Save the modified PDF
        nom_fitxer = f"Examen_{alumne['nom'].replace(' ', '_')}.pdf"
        path_sortida = os.path.join(output_dir_for_exams, nom_fitxer) # Use renamed parameter
        doc.save(path_sortida)
        doc.close()
        print(f"✅ Generated: {nom_fitxer}")

def create_zip_of_exams(exams_dir, output_dir_for_zip): # Renamed parameter for clarity
    """
    Creates a ZIP archive of all generated exams.
    """
    os.makedirs(output_dir_for_zip, exist_ok=True) # Use renamed parameter
    nom_fitxer_zip = os.path.join(output_dir_for_zip, "tots_els_examens") # Use renamed parameter # .zip will be added by make_archive

    if os.path.exists(exams_dir):
        print(f"📦 Compressing all PDFs from '{exams_dir}' to a ZIP archive...")
        shutil.make_archive(nom_fitxer_zip, 'zip', exams_dir)
        print(f"✅ ZIP created: {nom_fitxer_zip}.zip")
        return f"{nom_fitxer_zip}.zip"
    else:
        print(f"❌ ERROR: Exam directory '{exams_dir}' not found.")
        return None

if __name__ == "__main__":
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    base_pdf_template_path = os.path.join(current_script_dir, FITXER_BASE)
    
    # Corrected usage: OUTPUT_EXAMS_DIR and OUTPUT_ZIP_DIR are global configs
    output_exams_absolute_path = os.path.join(current_script_dir, OUTPUT_EXAMS_DIR)
    output_zip_absolute_path = os.path.join(current_script_dir, OUTPUT_ZIP_DIR)

    # Automatically discover CSV files in STUDENT_DATA_DIR
    # CORRECTED: student_data_full_path should be relative to current_script_dir
    student_data_full_path = os.path.join(current_script_dir, STUDENT_DATA_DIR) 
    
    os.makedirs(student_data_full_path, exist_ok=True) 
    csv_files_to_load = [os.path.join(student_data_full_path, f) for f in os.listdir(student_data_full_path) if f.lower().endswith('.csv')]

    alumnes_data = load_students_from_csv_files(csv_files_to_load) # Call the CSV loading function

    if not alumnes_data:
        print(f"❌ ERROR: No student data loaded from '{student_data_full_path}'. Exiting.", file=sys.stderr)
        sys.exit(1)

    try:
        generate_individual_exams(base_pdf_template_path, alumnes_data, output_exams_absolute_path) # Pass the absolute path
        zip_file_path = create_zip_of_exams(output_exams_absolute_path, output_zip_absolute_path) # Pass the absolute path
        
        if zip_file_path:
            print(f"Process completed. All exams generated and zipped to {zip_file_path}")
        else:
            print("Process completed with errors.")

    except Exception as e:
        print(f"An unexpected error occurred: {e}", file=sys.stderr)