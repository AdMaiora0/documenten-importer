import os
import shutil
import pandas as pd
from rich.console import Console
from rich.table import Table
from processor import DocumentProcessor

# Configuration
BASE_DIR = os.path.join(os.getcwd(), "test_data")
MAPPING_FILE = os.path.join(BASE_DIR, "mapping.xlsx")
SOURCE_DIR = os.path.join(BASE_DIR, "source_files")
OUTPUT_DIR = os.path.join(BASE_DIR, "output_files")
REPORT_FILE = os.path.join(BASE_DIR, "import_report.html")
SOURCE_COL = "Bestandsnaam"
TARGET_COL = "ClientID"

console = Console()

def run_test():
    console.print("[bold blue]--- Start Document Import Audit (Headless) ---[/bold blue]")
    
    # Clean output dir for fresh test
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR)

    processor = DocumentProcessor(MAPPING_FILE, SOURCE_DIR, OUTPUT_DIR, SOURCE_COL, TARGET_COL)
    
    # Run Process
    processor.process()
    
    # Generate Report
    processor.generate_report(REPORT_FILE)
    
    # CLI Output
    console.print("\n[bold]Resultaten:[/bold]")
    table = Table(title="Import Samenvatting")
    table.add_column("Metric", style="cyan")
    table.add_column("Waarde", style="magenta")
    
    table.add_row("Totaal Verwerkt", str(processor.stats["total"]))
    table.add_row("Succesvol", f"[green]{processor.stats['success']}[/green]")
    table.add_row("Mislukt", f"[red]{processor.stats['failed']}[/red]")
    table.add_row("Succespercentage", f"{processor.stats['success_rate']:.1f}%")
    console.print(table)

    # Test Zipping
    console.print("\n[bold yellow]Start Zipping Test...[/bold yellow]")
    # Use a smaller limit for testing splitting if needed, but user asked for >1GB data.
    # We generated ~1.5GB data. So 1GB limit should trigger split.
    zips = processor.create_zips(max_size_bytes=1024*1024*1024) 
    
    console.print(f"Zips created: {len(zips)}")
    for z in zips:
        size_mb = os.path.getsize(z) / (1024*1024)
        console.print(f"- {os.path.basename(z)} ({size_mb:.2f} MB)")

if __name__ == "__main__":
    run_test()
