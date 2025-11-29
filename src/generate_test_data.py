import pandas as pd
import os
import random
from datetime import datetime, timedelta

def generate_test_data():
    base_dir = os.path.join(os.getcwd(), "test_data")
    source_dir = os.path.join(base_dir, "source_files")
    output_dir = os.path.join(base_dir, "output_files")
    
    # Clean up and recreate directories
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
    
    if not os.path.exists(source_dir):
        os.makedirs(source_dir)
        
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Sample data generation
    data = []
    subjects = ["Verwijsbrief", "Huisartsenbrief", "Behandelovereenkomst", "Toestemmingsformulier", "Intakeverslag"]
    extensions = [".pdf", ".docx", ".png", ".jpg"]
    
    # Generate 1000 records
    # 500 Flat files (Scenario 1)
    # 500 Folders (Scenario 2)
    
    for i in range(1, 1001):
        # Introduce some data errors
        is_missing_client = random.random() < 0.01
        is_missing_filename = random.random() < 0.01
        
        client_id = random.randint(1, 6) if not is_missing_client else None
        dossier_id = 1
        subject = random.choice(subjects) if random.random() > 0.3 else "" 
        safe_subject = subject.replace(" ", "_") if subject else "document"
        
        date_val = datetime(2025, 5, 1) + timedelta(days=i)
        
        # Determine Scenario
        is_folder_scenario = i > 500
        
        if is_folder_scenario:
            # Scenario 2: Folder with files
            # Excel contains the Folder Name (or ID that matches folder)
            # Let's say Excel has "Dossier_<i>"
            folder_name = f"Dossier_{i}"
            excel_filename = folder_name if not is_missing_filename else None
            
            # Create folder on disk
            folder_path = os.path.join(source_dir, folder_name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                
            # Create 2-3 files inside
            for j in range(random.randint(1, 3)):
                sub_ext = random.choice(extensions)
                sub_filename = f"{safe_subject}_{j}{sub_ext}"
                sub_path = os.path.join(folder_path, sub_filename)
                with open(sub_path, "wb") as f:
                    f.write(os.urandom(1024 * 100)) # 100KB
                    
        else:
            # Scenario 1: Flat file
            # Excel contains "<i>_Subject.ext" OR just "<i>" if we want to test fuzzy matching?
            # User said: "waarin ergens in de documentnaam een document- of cliëntnummer staat"
            # Let's make the filename complex: "Scan_2025_Client_<i>_Subject.ext"
            # And Excel just has "<i>" or the full name?
            # Usually mapping files have the specific ID or Name.
            # Let's stick to: Excel has "<i>" (ID) and file is "<i>_Subject.ext" to test fuzzy match?
            # OR Excel has full filename.
            # Let's try to mix it up.
            # 1-250: Exact match
            # 251-500: Fuzzy match (Excel has ID, file has ID in name)
            
            ext = random.choice(extensions)
            
            if i <= 250:
                # Exact match
                filename = f"{i}_{safe_subject}{ext}"
                excel_filename = filename
            else:
                # Fuzzy match: Excel has ID "<i>", File is "Doc_<i>_v1.pdf"
                filename = f"Doc_{i}_v1{ext}"
                excel_filename = str(i)
            
            excel_filename = excel_filename if not is_missing_filename else None
            
            # Create file
            file_path = os.path.join(source_dir, filename)
            with open(file_path, "wb") as f:
                f.write(os.urandom(1024 * 500)) # 500KB

        data.append({
            "ID": i,
            "ClientID": client_id,
            "DossierID": dossier_id,
            "Onderwerp": subject,
            "Bestandsnaam": excel_filename,
            "Datumtijd": date_val
        })

    # Create DataFrame 


    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    excel_path = os.path.join(base_dir, "mapping.xlsx")
    df.to_excel(excel_path, index=False)
    
    print(f"Test data generated in {base_dir}")
    print(f"- Excel: {excel_path}")
    print(f"- Source Files: {source_dir}")
    print(f"- Output Directory: {output_dir}")

if __name__ == "__main__":
    generate_test_data()
