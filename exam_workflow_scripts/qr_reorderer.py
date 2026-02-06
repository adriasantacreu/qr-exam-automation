import fitz # PyMuPDF
from PIL import Image
import numpy as np
import cv2
import io
import os
import time
from collections import defaultdict
import json
import sys

# Import shared utility for loading student data
from exam_workflow_scripts.csv_utils import load_students_from_csv_files

# --- CONFIGURATION ---
DPI_DETECCIO = 300 # High quality for QR detection
DPI_SORTIDA = 150  # Output PDF DPI
QUALITAT_JPG = 80  # Quality for JPG conversion when inserting images into new PDF

# --- DIRECTORY PATHS (These should be configurable, possibly via environment variables or a config file) ---
# Input folder for raw scanned PDFs
CARPETA_ENTRADA = "scans_raw" 
# Output folder for reordered PDFs (for correction)
CARPETA_SORTIDA = "exams_for_correction"
# Debug folder for failed QR detections
CARPETA_DEBUG = "qr_debug_failures"
# Directory for student CSVs
STUDENT_DATA_DIR = "student_data"

# Page grouping for sorting (e.g., [[4], [5,6], ...])
# This defines the desired order of pages for correction (e.g., all page 1s, then all page 2s)
GRUPS_PAGINES = [[4], [5,6], [7,8], [9,10], [11,12]] 

# ==============================================================================
# 1. STUDENT DATA LOADING
# ==============================================================================
DICCIONARI_ALUMNES = {}
def carregar_base_dades_alumnes():
    global DICCIONARI_ALUMNES
    print("📊 Loading student data from CSV files...")
    
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    student_data_full_path = os.path.join(current_script_dir, STUDENT_DATA_DIR)
    os.makedirs(student_data_full_path, exist_ok=True) # Ensure student_data directory exists

    csv_files_to_load = [os.path.join(student_data_full_path, f) for f in os.listdir(student_data_full_path) if f.lower().endswith('.csv')]
    
    students_list = load_students_from_csv_files(csv_files_to_load) # Call CSV loading function
    
    if not students_list:
        print(f"❌ ERROR: No student data loaded from '{student_data_full_path}'. Cannot proceed with QR reordering.", file=sys.stderr)
        sys.exit(1) # Exit if no student data is found

    for student in students_list:
        DICCIONARI_ALUMNES[student["id"]] = student["nom"]
    print(f"✅ Loaded data for {len(DICCIONARI_ALUMNES)} students.")

# ==============================================================================
# 2. QR READER MOTORS (Ensure these libraries are installed via pip, e.g., 'pyzbar' and 'qreader')
# ==============================================================================
QREADER_DISPONIBLE = False
PYZBAR_DISPONIBLE = False

try:
    from qreader import QReader
    qreader_instance = QReader(model_size='s')
    QREADER_DISPONIBLE = True
    print("✅ QReader Activated")
except ImportError:
    print("⚠️ QReader not found. Install with `pip install qreader` for better QR detection.")
except Exception as e:
    print(f"❌ Error initializing QReader: {e}")

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    PYZBAR_DISPONIBLE = True
    print("✅ Pyzbar Activated")
except ImportError:
    print("⚠️ Pyzbar not found. Install with `pip install pyzbar` for QR detection.")
except Exception as e:
    print(f"❌ Error initializing Pyzbar: {e}")

if not QREADER_DISPONIBLE and not PYZBAR_DISPONIBLE:
    print("❌ No QR reader library available. Please install 'qreader' or 'pyzbar'.")
    sys.exit(1)


# ==============================================================================
# 3. IMAGE & ROTATION FUNCTIONS (CORE QR DETECTION LOGIC)
# ==============================================================================
def extreure_zona_qr(page):
    rect = page.rect
    marge = ZONA_QR["marge_seguretat"] if ZONA_QR["activat"] else 0
    if ZONA_QR["activat"]:
        x1 = max(0, ZONA_QR["x1"] - marge) * rect.width
        y1 = max(0, ZONA_QR["y1"] - marge) * rect.height
        x2 = min(1, ZONA_QR["x2"] + marge) * rect.width
        y2 = min(1, ZONA_QR["y2"] + marge) * rect.height
        clip = fitz.Rect(x1, y1, x2, y2)
        pix = page.get_pixmap(dpi=DPI_DETECCIO, clip=clip)
    else:
        pix = page.get_pixmap(dpi=DPI_DETECCIO)
    return np.array(Image.open(io.BytesIO(pix.tobytes("png"))))

def convertir_a_bgr(img_np):
    if len(img_np.shape) == 2: return cv2.cvtColor(img_np, cv2.COLOR_GRAY2BGR)
    elif img_np.shape[2] == 4: return cv2.cvtColor(img_np, cv2.COLOR_RGBA2BGR)
    elif img_np.shape[2] == 3: return cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
    return img_np

def rotar_imatge_graus(img, angle):
    """Rotates the image by X degrees, keeping the center."""
    if angle == 0: return img
    (h, w) = img.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    return cv2.warpAffine(img, M, (w, h), borderValue=(255, 255, 255))

def llegir_qr_bateria_proves(img_np):
    img_bgr = convertir_a_bgr(img_np)
    angles_a_provar = [0, 90, 180, 270, 5, -5, 10, -10] # Angles for robustness
    preprocessaments = [
        ("Normal", lambda x: x),
        ("Contrast", lambda x: cv2.convertScaleAbs(x, alpha=1.5, beta=0)),
        ("Binari", lambda x: cv2.cvtColor(cv2.adaptiveThreshold(cv2.cvtColor(x, cv2.COLOR_BGR2GRAY), 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2), cv2.COLOR_GRAY2BGR))
    ]

    for angle in angles_a_provar:
        img_rot = rotar_imatge_graus(img_bgr, angle)
        for nom_proc, func_proc in preprocessaments:
            try:
                img_final = func_proc(img_rot)

                # A. OpenCV
                det = cv2.QRCodeDetector()
                data, _, _ = det.detectAndDecode(img_final)
                if data: return data, f"OpenCV/{nom_proc}/{angle}º"

                # B. QReader
                if QREADER_DISPONIBLE:
                    res = qreader_instance.detect_and_decode(image=img_final)
                    if res and res[0]: return res[0], f"QReader/{nom_proc}/{angle}º"

                # C. Pyzbar
                if PYZBAR_DISPONIBLE and nom_proc in ["Normal", "Binari"]:
                    gray = cv2.cvtColor(img_final, cv2.COLOR_BGR2GRAY)
                    res = pyzbar_decode(gray)
                    if res: return res[0].data.decode("utf-8"), f"Pyzbar/{angle}º"
            except Exception:
                continue
    return None, None

def parsejar_qr(qr_text):
    if not qr_text: return None
    if "-" in qr_text:
        parts = qr_text.split("-")
        codi = parts[0].strip()
        try: num = int(parts[-1].strip())
        except ValueError: num = 999 # Handle cases where page number is not an int
        return {"id": codi, "pag": num}
    return None

def obtenir_index_grup(num_pagina):
    for index, grup in enumerate(GRUPS_PAGINES):
        if num_pagina in grup: return index
    return 999 # Default for pages not in any group

def guardar_debug(img_np, nom_pdf, i):
    os.makedirs(CARPETA_DEBUG, exist_ok=True) # Ensure debug folder exists
    cv2.imwrite(f"{CARPETA_DEBUG}/FAIL_{os.path.splitext(nom_pdf)[0]}_pag{i}.jpg", convertir_a_bgr(img_np))

# ==============================================================================
# 4. ADVANCED LOGIC (INTERPOLATION AND EXTRAPOLATION)
# ==============================================================================
def arreglar_forats_logicament(pages_info):
    print("🧠 Applying Logical Intelligence for missing QRs...")
    pages_info.sort(key=lambda x: x['idx']) # Sort by original PDF order
    canvis = 0

    # 1. INTERPOLATION (filling gaps in the middle)
    for i in range(1, len(pages_info) - 1):
        actual = pages_info[i]
        if actual['id'] is None:
            prev = pages_info[i-1]
            next_p = pages_info[i+1]
            if (prev['id'] and next_p['id'] and prev['id'] == next_p['id']):
                diff_pag = next_p['pag_num'] - prev['pag_num']
                diff_idx = next_p['idx'] - prev['idx']
                if diff_pag == diff_idx and diff_idx > 0: # Ensure valid logical sequence and avoid division by zero
                    pag_esperada = prev['pag_num'] + (actual['idx'] - prev['idx'])
                    actual['id'] = prev['id']
                    actual['pag_num'] = pag_esperada
                    actual['nom'] = DICCIONARI_ALUMNES.get(prev['id'], f"⚠️ Desconegut ({prev['id']})")
                    actual['metode'] = "✨Deducció(Mid)"
                    actual['sort'] = (obtenir_index_grup(pag_esperada), actual['nom'], pag_esperada)
                    print(f"💡 RECOVERED (Mid): Page {actual['idx']+1} is {prev['id']}-{pag_esperada}")
                    canvis += 1

    # 2. EXTRAPOLATION BACKWARD (recovering Page 1)
    for i in range(1, len(pages_info)):
        actual = pages_info[i]
        prev = pages_info[i-1]
        if actual['id'] is not None and prev['id'] is None:
            if actual['pag_num'] > 1:
                pag_anterior = actual['pag_num'] - 1
                prev['id'] = actual['id']
                prev['pag_num'] = pag_anterior
                prev['nom'] = DICCIONARI_ALUMNES.get(actual['id'], f"⚠️ Desconegut ({actual['id']})")
                prev['metode'] = "✨Deducció(Start)"
                prev['sort'] = (obtenir_index_grup(pag_anterior), prev['nom'], pag_anterior)
                print(f"💡 RECOVERED (Start): Page {prev['idx']+1} is {actual['id']}-{pag_anterior}")
                canvis += 1

    # 3. EXTRAPOLATION FORWARD (recovering last page)
    for i in range(0, len(pages_info) - 1):
        actual = pages_info[i]
        next_p = pages_info[i+1]
        if actual['id'] is not None and next_p['id'] is None:
            pag_seguent = actual['pag_num'] + 1
            # Add a check to prevent extrapolation into another student's exam if next-next page belongs to a different student
            is_safe_extrapolation = True
            if i + 2 < len(pages_info) and pages_info[i+2]['id'] is not None and pages_info[i+2]['id'] != actual['id']:
                 is_safe_extrapolation = False # Next-next page belongs to a different student, so don't extrapolate here

            if is_safe_extrapolation:
                next_p['id'] = actual['id']
                next_p['pag_num'] = pag_seguent
                next_p['nom'] = DICCIONARI_ALUMNES.get(actual['id'], f"⚠️ Desconegut ({actual['id']})")
                next_p['metode'] = "✨Deducció(End)"
                next_p['sort'] = (obtenir_index_grup(pag_seguent), next_p['nom'], pag_seguent)
                print(f"💡 RECOVERED (End): Page {next_p['idx']+1} is {actual['id']}-{pag_seguent}")
                canvis += 1
    
    if canvis == 0: print(" (No deduction needed)")
    return pages_info

def auditoria_final(pages_info):
    registre = defaultdict(set)
    for p in pages_info:
        if p['id'] and p['id'] != "ZZ_NoQR":
            registre[p['nom']].add(p['pag_num'])

    print("📋 --- AUDIT ---")
    alguna_alerta = False
    for alumne, pagines in registre.items():
        if not pagines: continue
        min_p, max_p = min(pagines), max(pagines)
        rang_esperat = set(range(min_p, max_p + 1))
        faltes = rang_esperat - pagines
        if faltes:
            print(f"⚠️ {alumne}: Missing pages {sorted(list(faltes))}")
            alguna_alerta = True
    if not alguna_alerta: print("✅ All good.")
    print("--------------------")

# ==============================================================================
# 5. MAIN PROCESSING LOOP FOR PDF REORDERING
# ==============================================================================
def processar_un_pdf(ruta_in, ruta_out, DICCIONARI_ALUMNES_LOCAL):
    # Make DICCIONARI_ALUMNES accessible within this function
    global DICCIONARI_ALUMNES
    DICCIONARI_ALUMNES = DICCIONARI_ALUMNES_LOCAL 
    
    try:
        doc = fitz.open(ruta_in)
    except Exception as e:
        print(f"❌ ERROR: Could not open PDF file '{ruta_in}': {e}", file=sys.stderr)
        return False

    nom_base = os.path.basename(ruta_in)
    print(f"📂 Processing: '{nom_base}'...")

    pages_info = []

    # PHASE 1: QR Reading
    for i, page in enumerate(doc):
        img_np = extreure_zona_qr(page)
        qr, met = llegir_qr_bateria_proves(img_np)

        info = {
            "idx": i, # Original index in the scanned PDF
            "id": None, "pag_num": 9999, "nom": "ZZ_NoQR", "metode": "❌ Error",
            "sort": (999, "ZZ_NoQR", 9999) # Tuple for sorting: (group_index, student_name, page_number)
        }

        if qr:
            dades = parsejar_qr(qr)
            if dades:
                nom_alumne = DICCIONARI_ALUMNES.get(dades["id"], f"⚠️ Desconegut ({dades['id']})")
                print(f"✅ Page {i+1:2d} -> {nom_alumne:<25} (P.{dades['pag']}) [{met}]")
                info.update({
                    "id": dades["id"], "pag_num": dades["pag"], "nom": nom_alumne, "metode": met,
                    "sort": (obtenir_index_grup(dades["pag"]), nom_alumne, dades["pag"])
                })
            else:
                print(f"❌ Page {i+1:2d} -> Unreadable QR: {qr}")
                guardar_debug(img_np, nom_base, i+1)
        else:
            print(f"❌ Page {i+1:2d} -> No QR found.")
            guardar_debug(img_np, nom_base, i+1)
            
        pages_info.append(info)

    # PHASE 2: Intelligent Deduction
    pages_info = arreglar_forats_logicament(pages_info)

    # PHASE 3: Reorder and prepare metadata
    print("🔄 Generating reordered PDF and saving MAP data...")
    pages_info.sort(key=lambda x: x["sort"])

    dades_per_guardar = []
    nou_doc = fitz.open() # New PDF document for reordered pages

    for nou_index, item in enumerate(pages_info):
        # 1. Build new PDF
        original_page_index = item["idx"]
        pix = doc[original_page_index].get_pixmap(dpi=DPI_SORTIDA)
        nova = nou_doc.new_page(width=pix.width, height=pix.height)
        nova.insert_image(nova.rect, stream=pix.tobytes("jpg", jpg_quality=QUALITAT_JPG))

        # 2. Save data for future use (the JSON map)
        dades_per_guardar.append({
            "nova_pagina_pdf": nou_index, # New page index (0, 1, 2, ...) in the reordered PDF
            "alumne_id": item["id"],
            "alumne_nom": item["nom"],
            "pagina_examen": item["pag_num"]
        })

    # Save reordered PDF
    os.makedirs(os.path.dirname(ruta_out), exist_ok=True) # Ensure output directory exists
    nou_doc.save(ruta_out)
    doc.close()
    nou_doc.close()

    # --- NEW PHASE: SAVE THE JSON MAP FILE ---
    ruta_json = ruta_out.replace(".pdf", ".json")
    with open(ruta_json, "w", encoding="utf-8") as f:
        json.dump(dades_per_guardar, f, indent=4, ensure_ascii=False)
    print(f"💾 Data map saved to: {os.path.basename(ruta_json)}")

    # PHASE 4: Audit
    auditoria_final(pages_info)
    print("✨ Done!")
    return True

def main():
    # Load student data from Excel files (or a local placeholder)
    carregar_base_dades_alumnes()

    os.makedirs(CARPETA_ENTRADA, exist_ok=True)
    os.makedirs(CARPETA_SORTIDA, exist_ok=True)
    os.makedirs(CARPETA_DEBUG, exist_ok=True)

    arxius = sorted([f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith('.pdf')])
    print(f"🚀 STARTING QR Reordering Process (Fine Rotation + Bidirectional Logic) for {len(arxius)} PDFs.")

    if not arxius:
        print(f"⚠️ No PDF files found in '{CARPETA_ENTRADA}'. Please upload scanned exams.")
        return

    for f in arxius:
        in_p = os.path.join(CARPETA_ENTRADA, f)
        out_p = os.path.join(CARPETA_SORTIDA, f)
        
        # Skip if output file already exists
        if os.path.exists(out_p):
            print(f"⏭️ Skipping '{f}' (already processed).")
            continue
        
        processar_un_pdf(in_p, out_p, DICCIONARI_ALUMNES) # Pass DICCIONARI_ALUMNES to the function

if __name__ == "__main__":
    main()
