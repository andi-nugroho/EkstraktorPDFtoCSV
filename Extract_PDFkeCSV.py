import os
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from rich.console import Console
from rich.table import Table
from rich.text import Text

console = Console()

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 Ekstrak PDF ke CSV")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        # Tema GUI
        self.style = ttk.Style("darkly")

        # Frame Utama
        self.frame = ttk.Frame(root, padding=20)
        self.frame.pack(expand=True, fill="both")

        # Label Judul
        self.label_title = ttk.Label(
            self.frame, text="📄 Ekstrak PDF ke CSV", font=("Arial", 16, "bold")
        )
        self.label_title.pack(pady=10)

        # Tombol Pilih File
        self.btn_pilih = ttk.Button(
            self.frame, text="📂 Pilih File PDF", bootstyle=PRIMARY, command=self.pilih_pdf
        )
        self.btn_pilih.pack(pady=10)

        # Label Drop Area
        self.label_drop = ttk.Label(
            self.frame, text="🚀 Drag & Drop file PDF ke sini", font=("Arial", 12, "italic")
        )
        self.label_drop.pack(pady=5)

        # Log Area
        self.log_text = tk.Text(self.frame, height=10, width=70, state="disabled", wrap="word")
        self.log_text.pack(pady=10)

        # Progress Bar
        self.progress = ttk.Progressbar(self.frame, mode="indeterminate", length=200)
        self.progress.pack(pady=10)

        # Bind Drag & Drop
        self.root.bind("<Control-o>", lambda event: self.pilih_pdf())

    def pilih_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.ekstrak_pdf_ke_csv(file_path)

    def ekstrak_pdf_ke_csv(self, pdf_path):
        self.progress.start()
        self.log_message("🔄 Memulai ekstraksi data dari PDF...")

        try:
            output_dir = os.path.dirname(pdf_path)
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            all_data = []

            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        all_data.extend(table[1:])

            if not all_data:
                messagebox.showerror("Error", "❌ Tidak ada data yang diekstrak dari PDF.")
                self.log_message("❌ Tidak ada data ditemukan dalam PDF.")
                return

            # Konversi ke DataFrame
            header = ["No", "NIM", "Nama", "Kehadiran", "UTS", "UAS", "Sikap", "Quiz", "Tugas Kelompok", "Tugas Besar", "Nilai Akhir", "Mutu"]
            df = pd.DataFrame(all_data, columns=header)

            # Simpan ke CSV dan Excel
            csv_output = os.path.join(output_dir, f"{pdf_name}.csv")
            xlsx_output = os.path.join(output_dir, f"{pdf_name}.xlsx")
            df.to_csv(csv_output, index=False)
            df.to_excel(xlsx_output, index=False)

            # Tampilkan hasil di terminal
            self.tampilkan_hasil_terminal(df)

            messagebox.showinfo("Sukses", f"✅ Data berhasil diekstrak ke:\n{csv_output}\n{xlsx_output}")
            self.log_message(f"✅ Ekstraksi berhasil! File disimpan: {csv_output}, {xlsx_output}")

        except Exception as e:
            messagebox.showerror("Error", f"⚠️ Gagal mengekstrak PDF: {e}")
            self.log_message(f"⚠️ Error: {e}")

        finally:
            self.progress.stop()

    def tampilkan_hasil_terminal(self, df):
        table = Table(title="📊 Hasil Ekstraksi PDF", show_lines=True)

        header_colors = {
            "No": "white",
            "NIM": "cyan",
            "Nama": "yellow",
            "Kehadiran": "green",
            "UTS": "blue",
            "UAS": "blue",
            "Sikap": "magenta",
            "Quiz": "cyan",
            "Tugas Kelompok": "bright_blue",
            "Tugas Besar": "red",
            "Nilai Akhir": "white",
            "Mutu": "bold red",
        }

        for col in df.columns:
            table.add_column(col, style=header_colors.get(col, "white"))

        for _, row in df.iterrows():
            row_values = []
            for col in df.columns:
                value = str(row[col])
                if col == "Mutu":
                    warna_mutu = {
                        "A": "bold green",
                        "B": "bold blue",
                        "C": "bold yellow",
                        "D": "bold red",
                        "E": "bold magenta",
                        "TL": "bold cyan"
                    }
                    row_values.append(Text(value, style=warna_mutu.get(value, "white")))
                else:
                    row_values.append(value)

            table.add_row(*row_values)

        console.print(table)

    def log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

if __name__ == "__main__":
    root = ttk.Window(themename="darkly")
    app = PDFExtractorApp(root)
    root.mainloop()
