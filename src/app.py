import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from processor import DocumentProcessor
import threading

class DocumentImporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Importer")
        self.root.geometry("700x600")

        # Variables
        self.mapping_file_path = tk.StringVar()
        self.source_dir_path = tk.StringVar()
        self.output_dir_path = tk.StringVar()
        self.source_col = tk.StringVar()
        self.target_col = tk.StringVar()
        self.processor = None

        # UI Layout
        self.create_widgets()

    def create_widgets(self):
        # Padding
        pad_opts = {'padx': 10, 'pady': 5}

        # 1. Select Mapping File
        grp_mapping = ttk.LabelFrame(self.root, text="1. Mapping Bestand (Excel)", padding=10)
        grp_mapping.pack(fill="x", **pad_opts)
        
        ttk.Entry(grp_mapping, textvariable=self.mapping_file_path).pack(side="left", fill="x", expand=True)
        ttk.Button(grp_mapping, text="Bladeren...", command=self.browse_mapping).pack(side="right", padx=(5, 0))

        # 2. Select Columns (Hidden initially)
        self.grp_columns = ttk.LabelFrame(self.root, text="2. Kolommen Selecteren", padding=10)
        
        ttk.Label(self.grp_columns, text="Bron Kolom (Bestandsnaam):").grid(row=0, column=0, sticky="w")
        self.cb_source = ttk.Combobox(self.grp_columns, textvariable=self.source_col, state="readonly")
        self.cb_source.grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Label(self.grp_columns, text="Doel Kolom (Cliëntnummer):").grid(row=1, column=0, sticky="w")
        self.cb_target = ttk.Combobox(self.grp_columns, textvariable=self.target_col, state="readonly")
        self.cb_target.grid(row=1, column=1, sticky="ew", padx=5)

        # 3. Select Directories
        grp_dirs = ttk.LabelFrame(self.root, text="3. Mappen Selecteren", padding=10)
        grp_dirs.pack(fill="x", **pad_opts)

        ttk.Label(grp_dirs, text="Bronmap (Bestanden):").pack(anchor="w")
        frame_src = ttk.Frame(grp_dirs)
        frame_src.pack(fill="x")
        ttk.Entry(frame_src, textvariable=self.source_dir_path).pack(side="left", fill="x", expand=True)
        ttk.Button(frame_src, text="Selecteer Map", command=lambda: self.browse_dir(self.source_dir_path)).pack(side="right", padx=(5, 0))

        ttk.Label(grp_dirs, text="Doelmap (Output):").pack(anchor="w", pady=(10, 0))
        frame_dst = ttk.Frame(grp_dirs)
        frame_dst.pack(fill="x")
        ttk.Entry(frame_dst, textvariable=self.output_dir_path).pack(side="left", fill="x", expand=True)
        ttk.Button(frame_dst, text="Selecteer Map", command=lambda: self.browse_dir(self.output_dir_path)).pack(side="right", padx=(5, 0))

        # 4. Action
        self.btn_run = ttk.Button(self.root, text="Start Verwerking", command=self.start_processing_thread, state="disabled")
        self.btn_run.pack(pady=20)

        # Progress
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=100, mode='determinate')
        self.progress.pack(fill="x", padx=20, pady=5)

        # Log area
        self.log_text = tk.Text(self.root, height=10, width=70)
        self.log_text.pack(padx=10, pady=5, fill="both", expand=True)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def browse_mapping(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.mapping_file_path.set(filename)
            try:
                import pandas as pd
                df = pd.read_excel(filename)
                columns = df.columns.tolist()
                self.cb_source['values'] = columns
                self.cb_target['values'] = columns
                
                # Try to guess columns
                for col in columns:
                    if "doc" in col.lower() or "bestand" in col.lower():
                        self.cb_source.set(col)
                    if "pat" in col.lower() or "cli" in col.lower() or "nr" in col.lower():
                        self.cb_target.set(col)

                self.grp_columns.pack(fill="x", padx=10, pady=5, after=self.root.children[list(self.root.children.keys())[0]])
                self.btn_run['state'] = 'normal'
                self.log(f"Bestand geladen: {len(df)} rijen gevonden.")
            except Exception as e:
                messagebox.showerror("Fout", f"Kan Excel bestand niet lezen:\n{e}")

    def browse_dir(self, var):
        dirname = filedialog.askdirectory()
        if dirname:
            var.set(dirname)

    def start_processing_thread(self):
        threading.Thread(target=self.run_process, daemon=True).start()

    def run_process(self):
        source_col = self.source_col.get()
        target_col = self.target_col.get()
        src_dir = self.source_dir_path.get()
        dst_dir = self.output_dir_path.get()
        mapping_file = self.mapping_file_path.get()

        if not all([source_col, target_col, src_dir, dst_dir, mapping_file]):
            messagebox.showwarning("Incompleet", "Vul alle velden in aub.")
            return

        self.btn_run['state'] = 'disabled'
        self.log("-" * 30)
        self.log("Start verwerking...")

        try:
            self.processor = DocumentProcessor(mapping_file, src_dir, dst_dir, source_col, target_col)
            
            def update_progress(current, total):
                self.progress['maximum'] = total
                self.progress['value'] = current
                self.root.update_idletasks()

            self.processor.process(progress_callback=update_progress)
            
            self.log(f"Verwerking klaar. Succes: {self.processor.stats['success']}, Mislukt: {self.processor.stats['failed']}")
            
            # Generate Report
            report_path = os.path.join(os.path.dirname(mapping_file), "import_report.html")
            self.processor.generate_report(report_path)
            self.log(f"Rapport gegenereerd: {report_path}")

            # Zip Output
            self.log("Bezig met zippen (dit kan even duren)...")
            self.progress['mode'] = 'indeterminate'
            self.progress.start()
            
            zips = self.processor.create_zips(max_size_bytes=1024*1024*1024) # 1GB
            
            self.progress.stop()
            self.progress['mode'] = 'determinate'
            self.progress['value'] = self.progress['maximum']
            
            self.log(f"Zips aangemaakt: {len(zips)}")
            for z in zips:
                self.log(f"- {os.path.basename(z)}")

            messagebox.showinfo("Klaar", f"Verwerking voltooid.\nSucces: {self.processor.stats['success']}\nZips: {len(zips)}")

        except Exception as e:
            self.log(f"CRITIQUE FOUT: {e}")
            messagebox.showerror("Fout", str(e))
        finally:
            self.btn_run['state'] = 'normal'

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentImporterApp(root)
    root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentImporterApp(root)
    root.mainloop()
