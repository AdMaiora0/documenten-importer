import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pathlib import Path

class DocumentImporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Importer")
        self.root.geometry("600x500")

        # Variables
        self.mapping_file_path = tk.StringVar()
        self.source_dir_path = tk.StringVar()
        self.output_dir_path = tk.StringVar()
        self.source_col = tk.StringVar()
        self.target_col = tk.StringVar()
        self.df = None

        # UI Layout
        self.create_widgets()

    def create_widgets(self):
        # Padding
        pad_opts = {'padx': 10, 'pady': 5}

        # 1. Select Mapping File
        grp_mapping = ttk.LabelFrame(self.root, text="1. Mapping Bestand (Excel)", padding=10)
        grp_mapping.pack(fill="x", **pad_opts)
        
        ttk.Entry(grp_mapping, textvariable=self.mapping_file_path, width=50).pack(side="left", fill="x", expand=True)
        ttk.Button(grp_mapping, text="Bladeren...", command=self.browse_mapping).pack(side="right")

        # 2. Select Columns (Hidden initially)
        self.grp_columns = ttk.LabelFrame(self.root, text="2. Kolommen Selecteren", padding=10)
        # Don't pack yet until file is loaded

        ttk.Label(self.grp_columns, text="Bron Kolom (Bestandsnaam/ID):").grid(row=0, column=0, sticky="w")
        self.cb_source = ttk.Combobox(self.grp_columns, textvariable=self.source_col, state="readonly")
        self.cb_source.grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Label(self.grp_columns, text="Doel Kolom (Cliëntnummer/Mapnaam):").grid(row=1, column=0, sticky="w")
        self.cb_target = ttk.Combobox(self.grp_columns, textvariable=self.target_col, state="readonly")
        self.cb_target.grid(row=1, column=1, sticky="ew", padx=5)

        # 3. Select Directories
        grp_dirs = ttk.LabelFrame(self.root, text="3. Mappen Selecteren", padding=10)
        grp_dirs.pack(fill="x", **pad_opts)

        ttk.Label(grp_dirs, text="Bronmap (waar de bestanden nu staan):").pack(anchor="w")
        frame_src = ttk.Frame(grp_dirs)
        frame_src.pack(fill="x")
        ttk.Entry(frame_src, textvariable=self.source_dir_path).pack(side="left", fill="x", expand=True)
        ttk.Button(frame_src, text="Selecteer Map", command=lambda: self.browse_dir(self.source_dir_path)).pack(side="right")

        ttk.Label(grp_dirs, text="Doelmap (waar de mappen gemaakt moeten worden):").pack(anchor="w", pady=(10, 0))
        frame_dst = ttk.Frame(grp_dirs)
        frame_dst.pack(fill="x")
        ttk.Entry(frame_dst, textvariable=self.output_dir_path).pack(side="left", fill="x", expand=True)
        ttk.Button(frame_dst, text="Selecteer Map", command=lambda: self.browse_dir(self.output_dir_path)).pack(side="right")

        # 4. Action
        self.btn_run = ttk.Button(self.root, text="Start Verwerking", command=self.run_process, state="disabled")
        self.btn_run.pack(pady=20)

        # Log area
        self.log_text = tk.Text(self.root, height=10, width=70)
        self.log_text.pack(padx=10, pady=5)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def browse_mapping(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.mapping_file_path.set(filename)
            try:
                self.df = pd.read_excel(filename)
                columns = self.df.columns.tolist()
                self.cb_source['values'] = columns
                self.cb_target['values'] = columns
                
                # Try to guess columns based on user description
                for col in columns:
                    if "doc" in col.lower() or "bestand" in col.lower():
                        self.cb_source.set(col)
                    if "pat" in col.lower() or "cli" in col.lower() or "nr" in col.lower():
                        self.cb_target.set(col)

                self.grp_columns.pack(fill="x", padx=10, pady=5, after=self.root.children[list(self.root.children.keys())[0]]) # Pack after first frame
                self.btn_run['state'] = 'normal'
                self.log(f"Bestand geladen: {len(self.df)} rijen gevonden.")
            except Exception as e:
                messagebox.showerror("Fout", f"Kan Excel bestand niet lezen:\n{e}")

    def browse_dir(self, var):
        dirname = filedialog.askdirectory()
        if dirname:
            var.set(dirname)

    def normalize_val(self, val):
        """Helper to handle Excel number/string formatting issues (e.g. 123.0 vs 123)"""
        s = str(val).strip()
        if s.endswith(".0"):
            return s[:-2]
        return s

    def run_process(self):
        source_col = self.source_col.get()
        target_col = self.target_col.get()
        src_dir = self.source_dir_path.get()
        dst_dir = self.output_dir_path.get()

        if not all([source_col, target_col, src_dir, dst_dir]):
            messagebox.showwarning("Incompleet", "Vul alle velden in aub.")
            return

        if self.df is None:
            return

        count_success = 0
        count_fail = 0

        self.log("-" * 30)
        self.log("Start verwerking...")
        self.log("Indexeren van bronbestanden...")

        # Pre-index source directory for performance and robustness
        # Maps 'filename_without_extension' -> list of full filenames
        source_files_map = {}
        try:
            for f in os.listdir(src_dir):
                name, ext = os.path.splitext(f)
                # Normalize the filename key as well just in case
                # But usually filenames are strings. 
                # We store the full filename to copy later.
                if name not in source_files_map:
                    source_files_map[name] = []
                source_files_map[name].append(f)
        except Exception as e:
            messagebox.showerror("Fout", f"Kan bronmap niet lezen: {e}")
            return

        for index, row in self.df.iterrows():
            # Normalize inputs (handle 123.0 from Excel)
            doc_id = self.normalize_val(row[source_col])
            client_id = self.normalize_val(row[target_col])

            # Skip empty rows
            if not doc_id or doc_id == 'nan' or not client_id or client_id == 'nan':
                continue

            # Create target folder
            target_path = os.path.join(dst_dir, client_id)
            try:
                os.makedirs(target_path, exist_ok=True)
            except Exception as e:
                self.log(f"FOUT: Kon map {target_path} niet maken: {e}")
                count_fail += 1
                continue

            # Find source files using the index
            # We look for files where the name (without extension) matches the doc_id
            found_files = source_files_map.get(doc_id, [])
            
            # Also check if there is a direct match for the doc_id as a file/folder name 
            # (e.g. if doc_id is "file.pdf" in excel, instead of just "file")
            if doc_id in os.listdir(src_dir):
                 if doc_id not in found_files:
                     found_files.append(doc_id)

            if found_files:
                for fname in found_files:
                    source_file_path = os.path.join(src_dir, fname)
                    try:
                        basename = os.path.basename(source_file_path)
                        dest_file_path = os.path.join(target_path, basename)
                        
                        if os.path.exists(dest_file_path):
                            self.log(f"LET OP: {basename} bestaat al in {client_id}, overgeslagen.")
                            # We don't count this as a hard fail, just a skip
                        else:
                            shutil.move(source_file_path, dest_file_path)
                            self.log(f"OK: {basename} -> {client_id}")
                            count_success += 1
                    except Exception as e:
                        self.log(f"FOUT: Kon {fname} niet verplaatsen: {e}")
                        count_fail += 1
            else:
                self.log(f"NIET GEVONDEN: Document '{doc_id}' niet gevonden.")
                count_fail += 1

        self.log("-" * 30)
        self.log(f"Klaar! {count_success} bestanden verplaatst.")
        messagebox.showinfo("Klaar", f"Verwerking voltooid.\nVerplaatst: {count_success}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentImporterApp(root)
    root.mainloop()
