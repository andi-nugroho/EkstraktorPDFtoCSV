import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os
from rich.console import Console
from rich.table import Table
from ttkbootstrap import Style
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
from pandastable import Table as PandasTable

console = Console()

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 Ekstrak PDF ke CSV & Excel")
        self.root.geometry("750x550")
        self.root.resizable(False, False)

        # Terapkan gaya modern
        self.style = Style(theme="superhero")
        self.current_theme = "superhero"

        # Frame utama
        self.main_frame = tk.Frame(root, bg="#1e1e1e")
        self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Label judul
        self.label_title = tk.Label(self.main_frame, text="📄 Ekstrak PDF ke CSV & Excel", font=("Arial", 18, "bold"), fg="white", bg="#1e1e1e")
        self.label_title.pack(pady=10)

        # Tombol pilih file
        self.btn_pilih = tk.Button(self.main_frame, text="📂 Pilih File PDF", font=("Arial", 12, "bold"),
                                   bg="#00A8E8", fg="white", padx=20, pady=10, relief="flat", command=self.pilih_pdf)
        self.btn_pilih.pack(pady=10)
        self.btn_pilih.bind("<Enter>", lambda e: self.btn_pilih.config(bg="#0077B6"))
        self.btn_pilih.bind("<Leave>", lambda e: self.btn_pilih.config(bg="#00A8E8"))

        # Opsi Simpan Format
        self.file_format = tk.StringVar(value="csv")
        frame_format = tk.Frame(self.main_frame, bg="#1e1e1e")
        frame_format.pack(pady=5)
        tk.Label(frame_format, text="Simpan sebagai:", fg="white", bg="#1e1e1e").pack(side="left", padx=5)
        tk.Radiobutton(frame_format, text="CSV", variable=self.file_format, value="csv", bg="#1e1e1e", fg="white").pack(side="left")
        tk.Radiobutton(frame_format, text="Excel (.xlsx)", variable=self.file_format, value="xlsx", bg="#1e1e1e", fg="white").pack(side="left")

        # Area log output
        self.log_frame = ScrolledFrame(self.main_frame, autohide=True, height=150, bootstyle="dark")
        self.log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text = tk.Text(self.log_frame, height=6, wrap="word", state="disabled", bg="#1e1e1e", fg="white", font=("Arial", 10))
        self.log_text.pack(fill="both", expand=True)

        # Tabel preview data
        self.table_frame = tk.Frame(self.main_frame, bg="#1e1e1e")
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Tombol Ganti Tema & Keluar
        self.footer_frame = tk.Frame(self.main_frame, bg="#1e1e1e")
        self.footer_frame.pack(pady=10)

        self.btn_ganti_tema = tk.Button(self.footer_frame, text="🌗 Ganti Tema", font=("Arial", 10), bg="#ffb703",
                                        fg="black", padx=10, command=self.ganti_tema)
        self.btn_ganti_tema.pack(side="left", padx=5)

        self.btn_keluar = tk.Button(self.footer_frame, text="❌ Keluar", font=("Arial", 10), bg="#e63946",
                                    fg="white", padx=10, command=root.quit)
        self.btn_keluar.pack(side="left", padx=5)

    def log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # Auto-scroll ke bawah
        self.log_text.config(state="disabled")
        console.print(f"[bold green]{message}[/bold green]")

    def pilih_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.ekstrak_pdf(file_path)

    def ekstrak_pdf(self, pdf_path):
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

            # Pilihan format penyimpanan
            file_format = self.file_format.get()
            if file_format == "csv":
                output_file = os.path.join(output_dir, f"{pdf_name}.csv")
                df.to_csv(output_file, index=False)
            else:
                output_file = os.path.join(output_dir, f"{pdf_name}.xlsx")
                df.to_excel(output_file, index=False)

            # Tampilkan tabel di terminal
            self.tampilkan_hasil_terminal(df)

            # Preview di GUI
            self.preview_data(df)

            messagebox.showinfo("Sukses", f"✅ Data berhasil diekstrak ke {output_file}")
            self.log_message(f"✅ Ekstraksi berhasil! File disimpan: {output_file}")

        except Exception as e:
            messagebox.showerror("Error", f"⚠️ Gagal mengekstrak PDF: {e}")
            self.log_message(f"⚠️ Error: {e}")

    def tampilkan_hasil_terminal(self, df):
        table = Table(title="📊 Hasil Ekstraksi PDF", show_lines=True)

        header_colors = {
            "No": "white", "NIM": "cyan", "Nama": "yellow",
            "Kehadiran": "green", "UTS": "blue", "UAS": "blue",
            "Sikap": "magenta", "Quiz": "cyan", "Tugas Kelompok": "bright_blue",
            "Tugas Besar": "red", "Nilai Akhir": "white", "Mutu": "bold red",
        }

        for col in df.columns:
            table.add_column(col, style=header_colors.get(col, "white"))

        for _, row in df.iterrows():
            table.add_row(*[str(val) for val in row])

        console.print(table)

    def preview_data(self, df):
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        table = PandasTable(self.table_frame, dataframe=df, showtoolbar=True, showstatusbar=True)
        table.show()

    def ganti_tema(self):
        self.current_theme = "darkly" if self.current_theme == "superhero" else "superhero"
        self.style.theme_use(self.current_theme)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
