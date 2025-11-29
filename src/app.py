import os
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from processor import DocumentProcessor
import threading

class DocumentImporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Importer")
        self.root.geometry("800x750")
        
        # Variables
        self.mapping_file_path = tk.StringVar()
        self.source_dir_path = tk.StringVar()
        # Default to Downloads folder
        default_output = os.path.join(os.path.expanduser("~"), "Downloads")
        self.output_dir_path = tk.StringVar(value=default_output)
        self.source_col = tk.StringVar()
        self.target_col = tk.StringVar()
        self.quarantine_var = tk.BooleanVar(value=False)
        self.processor = None

        # UI Layout
        self.create_widgets()

    def create_widgets(self):
        # Main Container
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)

        # Header
        lbl_header = ttk.Label(main_frame, text="Document Importer", font=("Segoe UI", 24, "bold"), bootstyle="primary")
        lbl_header.pack(pady=(0, 20))

        # 1. Configuration Section
        config_frame = ttk.Labelframe(main_frame, text="Configuratie", padding=15, bootstyle="info")
        config_frame.pack(fill=X, pady=10)

        # Mapping File
        ttk.Label(config_frame, text="Mapping Bestand (Excel):").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(config_frame, textvariable=self.mapping_file_path, width=50).grid(row=0, column=1, sticky=EW, padx=10, pady=5)
        ttk.Button(config_frame, text="Bladeren", command=self.browse_mapping, bootstyle="secondary-outline").grid(row=0, column=2, sticky=E, pady=5)

        # Columns (Initially disabled/empty)
        ttk.Label(config_frame, text="Bron Kolom (Bestandsnaam):").grid(row=1, column=0, sticky=W, pady=5)
        self.cb_source = ttk.Combobox(config_frame, textvariable=self.source_col, state="readonly", width=48)
        self.cb_source.grid(row=1, column=1, sticky=EW, padx=10, pady=5)

        ttk.Label(config_frame, text="Doel Kolom (Cliëntnummer):").grid(row=2, column=0, sticky=W, pady=5)
        self.cb_target = ttk.Combobox(config_frame, textvariable=self.target_col, state="readonly", width=48)
        self.cb_target.grid(row=2, column=1, sticky=EW, padx=10, pady=5)
        
        # Quarantine Checkbox
        ttk.Checkbutton(config_frame, text="Verplaats niet-gematchte bestanden naar _QUARANTINE", variable=self.quarantine_var, bootstyle="warning-round-toggle").grid(row=3, column=0, columnspan=2, sticky=W, pady=10)
        
        config_frame.columnconfigure(1, weight=1)

        # 2. Directories Section
        dir_frame = ttk.Labelframe(main_frame, text="Mappen", padding=15, bootstyle="info")
        dir_frame.pack(fill=X, pady=10)

        # Source Dir
        ttk.Label(dir_frame, text="Bronmap (Bestanden):").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(dir_frame, textvariable=self.source_dir_path).grid(row=0, column=1, sticky=EW, padx=10, pady=5)
        ttk.Button(dir_frame, text="Selecteer", command=lambda: self.browse_dir(self.source_dir_path), bootstyle="secondary-outline").grid(row=0, column=2, sticky=E, pady=5)

        # Output Dir
        ttk.Label(dir_frame, text="Doelmap (Output):").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(dir_frame, textvariable=self.output_dir_path).grid(row=1, column=1, sticky=EW, padx=10, pady=5)
        ttk.Button(dir_frame, text="Selecteer", command=lambda: self.browse_dir(self.output_dir_path), bootstyle="secondary-outline").grid(row=1, column=2, sticky=E, pady=5)
        
        dir_frame.columnconfigure(1, weight=1)

        # 3. Action Section
        action_frame = ttk.Frame(main_frame, padding=10)
        action_frame.pack(fill=X, pady=10)

        self.btn_run = ttk.Button(action_frame, text="Start Verwerking", command=self.start_processing_thread, state="disabled", bootstyle="success", width=20)
        self.btn_run.pack(side=LEFT)

        self.progress = ttk.Floodgauge(action_frame, bootstyle="success", font=("Segoe UI", 10), mask="{}%", orient=HORIZONTAL)
        self.progress.pack(side=LEFT, fill=X, expand=YES, padx=20)

        # Log Area
        self.log_text = ScrolledText(main_frame, height=10, autohide=True, bootstyle="round")
        self.log_text.pack(fill=BOTH, expand=YES, pady=10)
        self.log("Klaar voor gebruik. Selecteer eerst een mapping bestand.")

    def log(self, message):
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)

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
                source_candidates = ["bestandsnaam", "bestand", "document", "doc", "file"]
                target_candidates = ["clientnummer", "cliëntnummer", "clientid", "client_id", "patientnr", "nr"]

                def find_best_match(cols, candidates):
                    for candidate in candidates:
                        for col in cols:
                            if candidate in col.lower():
                                return col
                    return ""

                suggested_source = find_best_match(columns, source_candidates)
                if suggested_source:
                    self.cb_source.set(suggested_source)

                suggested_target = find_best_match(columns, target_candidates)
                if suggested_target:
                    self.cb_target.set(suggested_target)

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
        
        # Reset progress
        self.progress.configure(value=0, mask="{}%")
        self.progress.start()

        try:
            self.processor = DocumentProcessor(
                mapping_file, src_dir, dst_dir, source_col, target_col,
                quarantine=self.quarantine_var.get()
            )
            
            def update_progress(current, total):
                percent = (current / total) * 100
                self.progress.configure(value=percent, mask=f"{int(percent)}%")
                
            self.progress.stop() # Stop auto animation if any
            self.processor.process(progress_callback=update_progress)
            
            self.log(f"Verwerking klaar. Succes: {self.processor.stats['success']}, Mislukt: {self.processor.stats['failed']}")
            
            # Generate Report
            report_path = os.path.join(os.path.dirname(mapping_file), "import_report.html")
            self.processor.generate_report(report_path)
            self.log(f"Rapport gegenereerd: {report_path}")

            # Zip Output
            self.log("Bezig met zippen (dit kan even duren)...")
            self.progress.configure(mask="Zippen...", mode='indeterminate')
            self.progress.start()
            
            zips = self.processor.create_zips(max_size_bytes=1024*1024*1024) # 1GB
            
            self.progress.stop()
            self.progress.configure(value=100, mask="Klaar!", mode='determinate')
            
            self.log(f"Zips aangemaakt: {len(zips)}")
            for z in zips:
                self.log(f"- {os.path.basename(z)}")

            messagebox.showinfo("Klaar", f"Verwerking voltooid.\nSucces: {self.processor.stats['success']}\nZips: {len(zips)}")

        except Exception as e:
            self.log(f"CRITIQUE FOUT: {e}")
            messagebox.showerror("Fout", str(e))
        finally:
            self.btn_run['state'] = 'normal'
            self.progress.stop()

if __name__ == "__main__":
    # Create window with theme
    app = ttk.Window(themename="cosmo")
    DocumentImporterApp(app)
    app.mainloop()
