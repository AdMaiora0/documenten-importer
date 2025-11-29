import pandas as pd
import os
import random
from datetime import datetime, timedelta

def generate_test_data():
    base_dir = r"c:\Dev\documenten-importer\test_data"
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
    
    # Generate 20 records
    for i in range(1, 21):
        client_id = random.randint(1, 6) # Client IDs 1-6 like in screenshot
        dossier_id = 1
        subject = random.choice(subjects) if random.random() > 0.3 else "" # Some empty subjects
        
        # Construct filename similar to screenshot: "ID_Subject.ext"
        safe_subject = subject.replace(" ", "_") if subject else "document"
        ext = random.choice(extensions)
        filename = f"{i}_{safe_subject}{ext}"
        
        # Date
        date_val = datetime(2025, 5, 1) + timedelta(days=i)
        
        data.append({
            "ID": i,
            "ClientID": client_id,
            "DossierID": dossier_id,
            "Onderwerp": subject,
            "Bestandsnaam": filename,
            "Datumtijd": date_val
        })
        
        # Create the dummy file
        file_path = os.path.join(source_dir, filename)
        with open(file_path, "w") as f:
            f.write(f"Dummy content for {filename}")

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
