import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os
from rich.console import Console
from rich.table import Table
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk

console = Console()

class PDFExtractorApp:
    def __init__(self, root):
        # Daftar tema yang tersedia
        self.available_themes = [
            'superhero', 'darkly', 'solar', 'cyborg', 'vapor',
            'flatly', 'united', 'cosmo', 'journal', 'litera'
        ]
        self.current_theme_index = 0

        # Konfigurasi root window
        self.root = root
        self.root.title("📄 Ekstrak PDF ke CSV & Excel")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        # Atur ikon aplikasi
        self.set_app_icon()

        # Inisialisasi tema
        self.style = ttk.Style(theme=self.available_themes[self.current_theme_index])

        # Buat frame utama dengan padding dan border
        self.main_frame = ttk.Frame(root, padding="20 20 20 20")
        self.main_frame.pack(expand=True, fill="both")

        # Judul aplikasi dengan gaya modern
        self.create_title()

        # Tombol dan kontrol
        self.create_controls()

        # Frame log dengan scrollbar
        self.create_log_frame()

        # Frame pratinjau data
        self.create_preview_frame()

        # Footer dengan tombol aksi
        self.create_footer()

    def set_app_icon(self):
        try:
            if os.path.exists("icon.ico"):
                self.root.iconbitmap("icon.ico")
            elif os.path.exists("icon.png"):
                img = Image.open("icon.png")
                img = img.resize((32, 32), Image.LANCZOS)
                self.icon = ImageTk.PhotoImage(img)
                self.root.tk.call('wm', 'iconphoto', self.root._w, self.icon)
        except Exception as e:
            print(f"Error setting icon: {e}")

    def create_title(self):
        title_frame = ttk.Frame(self.main_frame)
        title_frame.pack(pady=10, fill='x')

        title_label = ttk.Label(
            title_frame,
            text="📄 Ekstrak PDF ke CSV & Excel",
            font=('-size', 18, '-weight', 'bold')
        )
        title_label.pack(expand=True)

    def create_controls(self):
        control_frame = ttk.Frame(self.main_frame)
        control_frame.pack(pady=10, fill='x')

        # Tombol Pilih File
        self.btn_pilih = ttk.Button(
            control_frame,
            text="📂 Pilih File PDF",
            command=self.pilih_pdf,
            style='primary.TButton'
        )
        self.btn_pilih.pack(side='left', padx=5)

        # Radio Button Format
        self.format_var = tk.StringVar(value="csv")
        format_frame = ttk.Frame(control_frame)
        format_frame.pack(side='left', padx=10)

        formats = [
            ("CSV", "csv"),
            ("Excel", "xlsx"),
            ("Keduanya", "both")
        ]

        for text, value in formats:
            rb = ttk.Radiobutton(
                format_frame,
                text=text,
                variable=self.format_var,
                value=value
            )
            rb.pack(side='left', padx=5)

    def create_log_frame(self):
        log_frame = ttk.LabelFrame(self.main_frame, text="Log Aktivitas")
        log_frame.pack(pady=10, fill='both', expand=True)

        self.log_text = tk.Text(
            log_frame,
            height=6,
            wrap='word',
            state='disabled'
        )
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)

    def create_preview_frame(self):
        preview_frame = ttk.LabelFrame(self.main_frame, text="Pratinjau Data")
        preview_frame.pack(pady=10, fill='both', expand=True)

        self.preview_text = tk.Text(
            preview_frame,
            height=8,
            wrap='none',
            state='disabled'
        )
        self.preview_text.pack(fill='both', expand=True, padx=5, pady=5)

    def create_footer(self):
        footer_frame = ttk.Frame(self.main_frame)
        footer_frame.pack(pady=10, fill='x')

        # Tombol Ganti Tema
        btn_ganti_tema = ttk.Button(
            footer_frame,
            text="🎨 Ganti Tema",
            command=self.ganti_tema,
            style='info.TButton'
        )
        btn_ganti_tema.pack(side='left', padx=5)

        # Tombol Keluar
        btn_keluar = ttk.Button(
            footer_frame,
            text="❌ Keluar",
            command=self.root.quit,
            style='danger.TButton'
        )
        btn_keluar.pack(side='right', padx=5)

    def ganti_tema(self):
        # Ganti tema secara berurutan
        self.current_theme_index = (self.current_theme_index + 1) % len(self.available_themes)
        new_theme = self.available_themes[self.current_theme_index]

        # Update tema
        self.style.theme_use(new_theme)
        self.log_message(f"Tema diganti menjadi: {new_theme}")

    def log_message(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
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

            header = ["No", "NIM", "Nama", "Kehadiran", "UTS", "UAS", "Sikap", "Quiz", "Tugas Kelompok", "Tugas Besar", "Nilai Akhir", "Mutu"]
            df = pd.DataFrame(all_data, columns=header)

            # Bersihkan data dari None atau string kosong
            df = df.dropna(how='all')
            df = df[df['NIM'].notna()]  # Pastikan kolom NIM tidak kosong

            format_terpilih = self.format_var.get()
            if format_terpilih in ["csv", "both"]:
                csv_output = os.path.join(output_dir, f"{pdf_name}.csv")
                df.to_csv(csv_output, index=False)
                self.log_message(f"✅ File CSV disimpan: {csv_output}")

            if format_terpilih in ["xlsx", "both"]:
                xlsx_output = os.path.join(output_dir, f"{pdf_name}.xlsx")
                df.to_excel(xlsx_output, index=False)
                self.log_message(f"✅ File Excel disimpan: {xlsx_output}")

            # Tampilkan pratinjau data
            self.preview_data(df)

            messagebox.showinfo("Sukses", "✅ Ekstraksi berhasil!")

        except Exception as e:
            messagebox.showerror("Error", f"⚠️ Gagal mengekstrak PDF: {e}")
            self.log_message(f"⚠️ Error: {e}")

    def preview_data(self, df):
        # Tampilkan pratinjau data di text widget
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', tk.END)
        self.preview_text.insert(tk.END, df.to_string())
        self.preview_text.config(state='disabled')

def main():
    root = ttk.Window(themename="superhero")
    app = PDFExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()