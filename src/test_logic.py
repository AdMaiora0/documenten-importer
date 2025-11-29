import os
import shutil
import pandas as pd

# Configuration for test
MAPPING_FILE = r"c:\Dev\documenten-importer\test_data\mapping.xlsx"
SOURCE_DIR = r"c:\Dev\documenten-importer\test_data\source_files"
OUTPUT_DIR = r"c:\Dev\documenten-importer\test_data\output_files"
SOURCE_COL = "Bestandsnaam"
TARGET_COL = "ClientID"

def normalize_val(val):
    s = str(val).strip()
    if s.endswith(".0"):
        return s[:-2]
    return s

def run_test():
    print("--- Start Headless Test ---")
    
    if not os.path.exists(MAPPING_FILE):
        print("Mapping file not found")
        return

    df = pd.read_excel(MAPPING_FILE)
    print(f"Loaded mapping with {len(df)} rows")

    # Indexing
    source_files_map = {}
    for f in os.listdir(SOURCE_DIR):
        name, ext = os.path.splitext(f)
        if name not in source_files_map:
            source_files_map[name] = []
        source_files_map[name].append(f)

    count_success = 0
    count_fail = 0

    for index, row in df.iterrows():
        doc_id = normalize_val(row[SOURCE_COL])
        client_id = normalize_val(row[TARGET_COL])

        if not doc_id or doc_id == 'nan' or not client_id or client_id == 'nan':
            continue

        target_path = os.path.join(OUTPUT_DIR, client_id)
        os.makedirs(target_path, exist_ok=True)

        found_files = source_files_map.get(doc_id, [])
        
        # Also check if there is a direct match for the doc_id as a file/folder name 
        if doc_id in os.listdir(SOURCE_DIR):
             if doc_id not in found_files:
                 found_files.append(doc_id)

        if found_files:
            for fname in found_files:
                source_file_path = os.path.join(SOURCE_DIR, fname)
                basename = os.path.basename(source_file_path)
                dest_file_path = os.path.join(target_path, basename)
                
                # COPY instead of MOVE for test preservation so you can still use the GUI
                shutil.copy(source_file_path, dest_file_path)
                print(f"OK: {basename} -> {client_id}")
                count_success += 1
        else:
            print(f"NOT FOUND: {doc_id}")
            count_fail += 1

    print(f"--- Done. Success: {count_success}, Fail: {count_fail} ---")

if __name__ == "__main__":
    run_test()
