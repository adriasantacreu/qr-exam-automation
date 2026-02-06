import csv
import os
import sys

# --- CSV Column Mapping ---
CSV_COL_ID_KEY = 'idalu' 
CSV_COL_NOM_KEY = 'nom'
CSV_COL_EMAIL_KEY = 'mail'

def load_students_from_csv_files(csv_file_paths): # Accepts a list of file paths
    all_students = []
    for csv_file in csv_file_paths: # Iterate through provided file paths
        if not os.path.exists(csv_file):
            print(f"⚠️ Warning: CSV file '{csv_file}' not found. Skipping.", file=sys.stderr)
            continue
        try:
            with open(csv_file, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                
                # Convert headers to lowercase for case-insensitive matching
                if reader.fieldnames:
                    fieldnames = [str(header).lower() for header in reader.fieldnames]
                    reader.fieldnames = fieldnames # Update fieldnames for DictReader
                
                # Find column indices (or check for existence) for 'idalu', 'nom', 'mail'
                if CSV_COL_ID_KEY not in fieldnames or \
                   CSV_COL_NOM_KEY not in fieldnames or \
                   CSV_COL_EMAIL_KEY not in fieldnames:
                    print(f"❌ Error: Could not find '{CSV_COL_ID_KEY}', '{CSV_COL_NOM_KEY}', or '{CSV_COL_EMAIL_KEY}' columns (case-insensitive) in '{csv_file}'. Skipping.", file=sys.stderr)
                    continue

                for row_idx, row in enumerate(reader, 2): # row_idx starts from 2 for data rows after header
                    # Access data using lowercase keys
                    student_id = row.get(CSV_COL_ID_KEY)
                    student_nom = row.get(CSV_COL_NOM_KEY)
                    student_email = row.get(CSV_COL_EMAIL_KEY)

                    if student_id and student_nom and student_email:
                        all_students.append({
                            "id": str(student_id).strip(),
                            "nom": str(student_nom).strip(),
                            "email": str(student_email).strip()
                        })
                    else:
                        print(f"⚠️ Warning: Skipping row {row_idx} in '{csv_file}' due to missing data in required columns.", file=sys.stderr)

        except Exception as e:
            print(f"❌ Error reading CSV file '{csv_file}': {e}", file=sys.stderr)
    return all_students
