import os
import shutil
import pandas as pd
import zipfile
from datetime import datetime
from jinja2 import Template

class DocumentProcessor:
    def __init__(self, mapping_file, source_dir, output_dir, source_col, target_col):
        self.mapping_file = mapping_file
        self.source_dir = source_dir
        self.output_dir = output_dir
        self.source_col = source_col
        self.target_col = target_col
        self.audit_log = []
        self.stats = {
            "total": 0,
            "success": 0,
            "failed": 0,
            "errors": {},
            "client_counts": {}
        }

    def normalize_val(self, val):
        s = str(val).strip()
        if s.lower() == 'nan' or s.lower() == 'none' or not s:
            return None
        if s.endswith(".0"):
            return s[:-2]
        return s

    def process(self, progress_callback=None):
        if not os.path.exists(self.mapping_file):
            raise FileNotFoundError("Mapping file not found")

        df = pd.read_excel(self.mapping_file)
        
        # Index files
        files_on_disk = set(os.listdir(self.source_dir))

        total_rows = len(df)
        
        for index, row in df.iterrows():
            if progress_callback:
                progress_callback(index + 1, total_rows)

            self.stats["total"] += 1
            row_id = row.get("ID", index + 1)
            raw_doc_name = row.get(self.source_col)
            raw_client_id = row.get(self.target_col)
            
            doc_name = self.normalize_val(raw_doc_name)
            client_id = self.normalize_val(raw_client_id)
            
            log_entry = {
                "id": row_id,
                "filename": str(doc_name) if doc_name else "N/A",
                "client_id": str(client_id) if client_id else "N/A",
                "status": "PENDING",
                "message": ""
            }

            if not doc_name:
                self._log_error(log_entry, "Bestandsnaam ontbreekt in Excel")
                continue

            if not client_id:
                self._log_error(log_entry, "Cliënt ID ontbreekt in Excel")
                continue

            if doc_name not in files_on_disk:
                self._log_error(log_entry, f"Bestand niet gevonden: {doc_name}")
                continue

            try:
                target_path = os.path.join(self.output_dir, client_id)
                os.makedirs(target_path, exist_ok=True)
                
                src_file = os.path.join(self.source_dir, doc_name)
                dst_file = os.path.join(target_path, doc_name)
                
                if os.path.exists(dst_file):
                     log_entry["status"] = "SKIPPED"
                     log_entry["message"] = f"Bestand bestaat al in {client_id}"
                     # Not a failure, but not a fresh success either. Let's count as success for now or separate metric?
                     # User asked for "Wat is goed gegaan", "Wat is fout gegaan". Skipped is kinda success.
                     self.stats["success"] += 1
                else:
                    shutil.copy2(src_file, dst_file)
                    log_entry["status"] = "SUCCESS"
                    log_entry["message"] = f"Gekopieerd naar {client_id}"
                    self.stats["success"] += 1
                
                self.stats["client_counts"][client_id] = self.stats["client_counts"].get(client_id, 0) + 1
                
            except Exception as e:
                self._log_error(log_entry, f"Systeemfout: {str(e)}")

            self.audit_log.append(log_entry)

        self.stats["success_rate"] = (self.stats["success"] / self.stats["total"] * 100) if self.stats["total"] > 0 else 0
        self.stats["top_clients"] = sorted(self.stats["client_counts"].items(), key=lambda item: item[1], reverse=True)[:10]

    def _log_error(self, entry, message):
        entry["status"] = "ERROR"
        entry["message"] = message
        self.stats["errors"][message] = self.stats["errors"].get(message, 0) + 1
        self.stats["failed"] += 1
        self.audit_log.append(entry)

    def generate_report(self, report_path):
        template_str = """
        <!DOCTYPE html>
        <html lang="nl">
        <head>
            <meta charset="UTF-8">
            <title>Document Import Rapport</title>
            <style>
                body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f4f4f9; }
                h1, h2 { color: #333; }
                .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); border-radius: 8px; }
                .summary-box { display: flex; gap: 20px; margin-bottom: 20px; }
                .card { flex: 1; padding: 15px; border-radius: 5px; color: white; text-align: center; }
                .bg-blue { background-color: #007bff; }
                .bg-green { background-color: #28a745; }
                .bg-red { background-color: #dc3545; }
                .bg-orange { background-color: #fd7e14; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th, td { padding: 10px; border: 1px solid #ddd; text-align: left; }
                th { background-color: #f8f9fa; }
                tr:nth-child(even) { background-color: #f9f9f9; }
                .status-success { color: green; font-weight: bold; }
                .status-error { color: red; font-weight: bold; }
                .status-skipped { color: orange; font-weight: bold; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Document Import Audit Rapport</h1>
                <p>Gegenereerd op: {{ timestamp }}</p>

                <h2>Samenvatting</h2>
                <div class="summary-box">
                    <div class="card bg-blue">
                        <h3>{{ summary.total }}</h3>
                        <p>Totaal Verwerkt</p>
                    </div>
                    <div class="card bg-green">
                        <h3>{{ summary.success }}</h3>
                        <p>Succesvol</p>
                    </div>
                    <div class="card bg-red">
                        <h3>{{ summary.failed }}</h3>
                        <p>Mislukt</p>
                    </div>
                    <div class="card bg-orange">
                        <h3>{{ "%.1f"|format(summary.success_rate) }}%</h3>
                        <p>Succespercentage</p>
                    </div>
                </div>

                <h2>Foutanalyse</h2>
                {% if summary.error_counts %}
                <table>
                    <thead>
                        <tr>
                            <th>Foutmelding</th>
                            <th>Aantal</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for error, count in summary.error_counts.items() %}
                        <tr>
                            <td>{{ error }}</td>
                            <td>{{ count }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                {% else %}
                <p>Geen fouten gevonden.</p>
                {% endif %}

                <h2>Audit Log (Details)</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Rij ID</th>
                            <th>Bestandsnaam</th>
                            <th>Cliënt ID</th>
                            <th>Status</th>
                            <th>Opmerking</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for log in audit_log %}
                        <tr>
                            <td>{{ log.id }}</td>
                            <td>{{ log.filename }}</td>
                            <td>{{ log.client_id }}</td>
                            <td class="{{ 'status-success' if log.status == 'SUCCESS' else ('status-error' if log.status == 'ERROR' else 'status-skipped') }}">{{ log.status }}</td>
                            <td>{{ log.message }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </body>
        </html>
        """
        
        template = Template(template_str)
        html_content = template.render(
            timestamp=datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
            summary=self.stats,
            audit_log=self.audit_log
        )
        
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        return report_path

    def create_zips(self, max_size_bytes=1024*1024*1024): # 1GB default
        # Get all client folders
        client_folders = [f for f in os.listdir(self.output_dir) if os.path.isdir(os.path.join(self.output_dir, f))]
        
        # Sort numerically if possible, else string sort
        def sort_key(x):
            try:
                return int(x)
            except ValueError:
                return x
        
        client_folders.sort(key=sort_key)
        
        current_batch = []
        current_batch_size = 0
        zip_files_created = []

        for client_id in client_folders:
            client_path = os.path.join(self.output_dir, client_id)
            client_size = self._get_dir_size(client_path)
            
            # If adding this client exceeds max size AND we have a batch, zip the current batch
            if current_batch and (current_batch_size + client_size > max_size_bytes):
                zip_path = self._zip_batch(current_batch)
                zip_files_created.append(zip_path)
                current_batch = []
                current_batch_size = 0
            
            current_batch.append(client_id)
            current_batch_size += client_size
            
        # Zip remaining
        if current_batch:
            zip_path = self._zip_batch(current_batch)
            zip_files_created.append(zip_path)
            
        return zip_files_created

    def _get_dir_size(self, path):
        total = 0
        for dirpath, dirnames, filenames in os.walk(path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if not os.path.islink(fp):
                    total += os.path.getsize(fp)
        return total

    def _zip_batch(self, client_ids):
        if not client_ids:
            return None
            
        # Determine name range
        first = client_ids[0]
        last = client_ids[-1]
        
        # If numeric, we can make it look nice
        zip_name = f"Export_Clients_{first}_to_{last}.zip"
        zip_path = os.path.join(self.output_dir, zip_name) # Wait, user said output to output dir? 
        # "Ik wil dat je de output vervolgens zipt." 
        # Usually zips are placed alongside the output folder or inside it?
        # If I place it INSIDE the output dir, it might get recursively zipped if I'm not careful, 
        # but I am iterating over client folders specifically.
        # However, if I run this multiple times, I might zip old zips?
        # Let's place zips in the PARENT of output_dir or a specific 'zips' folder?
        # User said: "Als de output zip groter is dan 1gb...".
        # Let's put zips in the output_dir for now but ensure we only zip directories.
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for client_id in client_ids:
                client_path = os.path.join(self.output_dir, client_id)
                for root, dirs, files in os.walk(client_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        # Archive name should be relative to output_dir so it contains the client folder
                        arcname = os.path.relpath(file_path, self.output_dir)
                        zipf.write(file_path, arcname)
        
        return zip_path
