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
from PIL import Image, ImageTk

console = Console()

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 Ekstrak PDF ke CSV & Excel")
        self.root.geometry("900x600")
        self.root.minsize(700, 500)

        # Tambahkan fitur Restore Down / Maximize
        self.root.state('normal')

        # Ubah ikon aplikasi (Gunakan PNG sebagai alternatif)
        try:
            icon_image = Image.open("lampu.png")
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.root.iconphoto(False, icon_photo)
        except Exception as e:
            console.print(f"[bold red]⚠️ Ikon tidak ditemukan atau gagal dimuat: {e}[/bold red]")

        # Terapkan gaya modern
        self.style = Style(theme="superhero")

        # Frame utama
        self.main_frame = tk.Frame(root, bg="#1e1e1e")
        self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Label judul
        self.label_title = tk.Label(self.main_frame, text="📄 Ekstrak PDF ke CSV & Excel",
                                    font=("Arial", 18, "bold"), fg="white", bg="#1e1e1e")
        self.label_title.pack(pady=10)

        # Tombol pilih file
        self.btn_pilih = tk.Button(self.main_frame, text="📂 Pilih File PDF", font=("Arial", 12, "bold"),
                                   bg="#00A8E8", fg="white", padx=20, pady=10, relief="flat",
                                   command=self.pilih_pdf)
        self.btn_pilih.pack(pady=10)
        self.btn_pilih.bind("<Enter>", lambda e: self.btn_pilih.config(bg="#0077B6"))
        self.btn_pilih.bind("<Leave>", lambda e: self.btn_pilih.config(bg="#00A8E8"))

        # Opsi simpan file (diperbaiki tata letaknya)
        self.format_var = tk.StringVar(value="csv")
        self.frame_format = tk.Frame(self.main_frame, bg="#1e1e1e")
        self.frame_format.pack(pady=5)

        tk.Label(self.frame_format, text="Simpan sebagai:", fg="white", bg="#1e1e1e", font=("Arial", 11)).grid(row=0, column=0, columnspan=3)

        self.radio_csv = tk.Radiobutton(self.frame_format, text="CSV", variable=self.format_var, value="csv", fg="white", bg="#1e1e1e", selectcolor="#1e1e1e")
        self.radio_csv.grid(row=1, column=0, padx=10)

        self.radio_xlsx = tk.Radiobutton(self.frame_format, text="Excel (.xlsx)", variable=self.format_var, value="xlsx", fg="white", bg="#1e1e1e", selectcolor="#1e1e1e")
        self.radio_xlsx.grid(row=1, column=1, padx=10)

        self.radio_both = tk.Radiobutton(self.frame_format, text="CSV & Excel", variable=self.format_var, value="both", fg="white", bg="#1e1e1e", selectcolor="#1e1e1e")
        self.radio_both.grid(row=1, column=2, padx=10)

        # Area log output
        self.log_frame = ScrolledFrame(self.main_frame, autohide=True, height=100, bootstyle="dark")
        self.log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text = tk.Text(self.log_frame, height=5, wrap="word", state="disabled", bg="#1e1e1e", fg="white", font=("Arial", 10))
        self.log_text.pack(fill="both", expand=True)

        # Tabel preview data
        self.table_frame = tk.Frame(self.main_frame, bg="#1e1e1e")
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Tombol tambahan
        self.footer_frame = tk.Frame(self.main_frame, bg="#1e1e1e")
        self.footer_frame.pack(fill="x", pady=10)

        self.btn_ganti_tema = tk.Button(self.footer_frame, text="🎨 Ganti Tema", font=("Arial", 10, "bold"),
                                        bg="#0077B6", fg="white", padx=10, pady=5, relief="flat",
                                        command=self.ganti_tema)
        self.btn_ganti_tema.pack(side="left", padx=5)

        self.btn_keluar = tk.Button(self.footer_frame, text="❌ Keluar", font=("Arial", 10, "bold"),
                                    bg="#D9534F", fg="white", padx=10, pady=5, relief="flat",
                                    command=root.quit)
        self.btn_keluar.pack(side="right", padx=5)

    def log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
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

            header = ["No", "NIM", "Nama", "Kehadiran", "UTS", "UAS", "Sikap", "Quiz", "Tugas Kelompok", "Tugas Besar", "Nilai Akhir", "Mutu"]
            df = pd.DataFrame(all_data, columns=header)

            # Simpan file sesuai pilihan format
            format_terpilih = self.format_var.get()
            if format_terpilih in ["csv", "both"]:
                csv_output = os.path.join(output_dir, f"{pdf_name}.csv")
                df.to_csv(csv_output, index=False)
                self.log_message(f"✅ File CSV disimpan: {csv_output}")

            if format_terpilih in ["xlsx", "both"]:
                xlsx_output = os.path.join(output_dir, f"{pdf_name}.xlsx")
                df.to_excel(xlsx_output, index=False)
                self.log_message(f"✅ File Excel disimpan: {xlsx_output}")

            self.tampilkan_hasil_terminal(df)
            self.preview_data(df)

            messagebox.showinfo("Sukses", "✅ Ekstraksi berhasil!")
        except Exception as e:
            messagebox.showerror("Error", f"⚠️ Gagal mengekstrak PDF: {e}")
            self.log_message(f"⚠️ Error: {e}")

    def tampilkan_hasil_terminal(self, df):
        table = Table(title="📊 Hasil Ekstraksi PDF", show_lines=True)
        for col in df.columns:
            table.add_column(col, style="cyan")
        for _, row in df.iterrows():
            table.add_row(*[str(val) for val in row])
        console.print(table)

    def preview_data(self, df):
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        table = PandasTable(self.table_frame, dataframe=df, showtoolbar=True, showstatusbar=True)
        table.show()

    def ganti_tema(self):
        tema_baru = "darkly" if self.style.theme == "superhero" else "superhero"
        self.style.theme_use(tema_baru)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
