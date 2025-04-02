import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from ttkthemes import ThemedTk
from pandastable import Table
import fitz  # PyMuPDF

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ekstrak PDF ke CSV")
        self.root.geometry("900x600")

        # Tema awal
        self.current_theme = "equilux"
        self.root.set_theme(self.current_theme)

        # Frame utama
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        # Judul
        self.label_title = ttk.Label(main_frame, text="📄 Ekstrak PDF ke CSV", font=("Arial", 16, "bold"))
        self.label_title.pack(pady=10)

        # Tombol Pilih File
        self.btn_pilih_file = ttk.Button(main_frame, text="📂 Pilih File PDF", command=self.load_pdf)
        self.btn_pilih_file.pack(pady=5)

        # Log
        self.text_log = tk.Text(main_frame, height=5, state="disabled", wrap="word")
        self.text_log.pack(fill="both", expand=True, pady=5)

        # Frame tabel
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill="both", expand=True)

        # Tabel data
        self.data_table = Table(table_frame)
        self.data_table.show()

        # Frame tombol aksi
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        # Tombol Simpan CSV
        self.btn_simpan_csv = ttk.Button(button_frame, text="💾 Simpan CSV", command=lambda: self.save_file("csv"))
        self.btn_simpan_csv.grid(row=0, column=0, padx=5)

        # Tombol Simpan Excel
        self.btn_simpan_xlsx = ttk.Button(button_frame, text="📊 Simpan Excel", command=lambda: self.save_file("xlsx"))
        self.btn_simpan_xlsx.grid(row=0, column=1, padx=5)

        # Tombol Ganti Tema
        self.btn_ganti_tema = ttk.Button(button_frame, text="🌙 Ganti Tema", command=self.toggle_theme)
        self.btn_ganti_tema.grid(row=0, column=2, padx=5)

        # Tombol Keluar
        self.btn_keluar = ttk.Button(button_frame, text="❌ Keluar", command=self.root.quit)
        self.btn_keluar.grid(row=0, column=3, padx=5)

    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return

        self.log_message(f"Memulai ekstraksi data dari PDF: {file_path}")

        try:
            extracted_data = self.extract_text_from_pdf(file_path)
            df = pd.DataFrame(extracted_data)
            self.data_table.model.df = df
            self.data_table.redraw()

            self.log_message("✅ Ekstraksi berhasil!")

        except Exception as e:
            self.log_message(f"❌ Terjadi kesalahan: {str(e)}")

    def extract_text_from_pdf(self, file_path):
        extracted_data = []
        with fitz.open(file_path) as doc:
            for page in doc:
                text = page.get_text("text").split("\n")
                extracted_data.append(text)
        return extracted_data

    def save_file(self, file_type):
        file_path = filedialog.asksaveasfilename(defaultextension=f".{file_type}",
                                                 filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")])
        if not file_path:
            return

        try:
            df = self.data_table.model.df
            if file_type == "csv":
                df.to_csv(file_path, index=False)
            else:
                df.to_excel(file_path, index=False, engine="openpyxl")

            self.log_message(f"✅ File disimpan: {file_path}")

        except Exception as e:
            self.log_message(f"❌ Gagal menyimpan: {str(e)}")

    def log_message(self, message):
        self.text_log.config(state="normal")
        self.text_log.insert("end", message + "\n")
        self.text_log.config(state="disabled")

    def toggle_theme(self):
        if self.current_theme == "equilux":
            self.current_theme = "plastik"
        else:
            self.current_theme = "equilux"

        self.root.set_theme(self.current_theme)

# Jalankan aplikasi
if __name__ == "__main__":
    root = ThemedTk(theme="equilux")
    app = PDFExtractorApp(root)
    root.mainloop()
