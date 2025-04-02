import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os
from rich.console import Console
from rich.table import Table
from ttkbootstrap import Style
from ttkbootstrap.scrolled import ScrolledFrame
from pandastable import Table as PandasTable
from PIL import Image, ImageTk
from tkinter import ttk

console = Console()

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 Ekstrak PDF ke CSV & Excel")
        self.root.geometry("900x600")
        self.root.minsize(700, 500)

        # application icon
        self.set_icon()

        # modern style
        self.style = Style(theme="superhero")
        self.style = Style(theme="darkly")

        # Create main frame
        self.main_frame = self.create_main_frame()
        self.create_widgets()

    def set_icon(self):
        """Set the application icon."""
        if os.path.exists("icon.ico"):
            self.root.iconbitmap("icon.ico")
        elif os.path.exists("UMC-LOGO.png"):
            img = Image.open("UMC-LOGO.png")
            img = img.resize((32, 32), Image.LANCZOS)
            self.icon = ImageTk.PhotoImage(img)
            self.root.tk.call('wm', 'iconphoto', self.root._w, self.icon)

    def create_main_frame(self):
        """Create and return the main frame."""
        main_frame = tk.Frame(self.root, bg="#1e1e1e")
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)
        return main_frame

    def create_widgets(self):
        """Create and place all widgets in the main frame."""
        self.create_title_label()
        self.create_file_selection_button()
        self.create_format_selection()
        self.create_log_output_area()
        self.create_data_preview_table()
        self.create_footer_buttons()

    def create_title_label(self):
        """Create and place the title label."""
        label_title = tk.Label(self.main_frame, text="📄 Ekstrak PDF ke CSV & Excel",
                                font=("Arial", 18, "bold"), fg="white", bg="#1e1e1e")
        label_title.pack(pady=10)

    def create_file_selection_button(self):
        """Create and place the file selection button."""
        self.btn_pilih = tk.Button(self.main_frame, text="📂 Pilih File PDF", font=("Arial", 12, "bold"),
                                   bg="#00A8E8", fg="white", padx=20, pady=10, relief="flat",
                                   command=self.pilih_pdf)
        self.btn_pilih.pack(pady=10)
        self.btn_pilih.bind("<Enter>", lambda e: self.btn_pilih.config(bg="#0077B6"))
        self.btn_pilih.bind("<Leave>", lambda e: self.btn_pilih.config(bg="#00A8E8"))

    def create_format_selection(self):
        """Create and place the format selection options."""
        self.format_var = tk.StringVar(value="csv")
        format_frame = tk.Frame(self.main_frame, bg="#1e1e1e")
        format_frame.pack(pady=5)

        # Radio buttons for format selection
        formats = [("CSV", "csv"), ("Excel (.xlsx)", "xlsx"), ("CSV & Excel", "both")]
        for idx, (text, value) in enumerate(formats):
            radio = tk.Radiobutton(format_frame, text=text, variable=self.format_var, value=value,
                                   fg="white", bg="#1e1e1e", selectcolor="#1e1e1e")
            radio.grid(row=0, column=idx, padx=10)

    def create_log_output_area(self):
        """Create and place the log output area."""
        self.log_frame = ScrolledFrame(self.main_frame, autohide=True, height=100, bootstyle="dark")
        self.log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text = tk.Text(self.log_frame, height=5, wrap="word", state="disabled", bg="#1e1e1e", fg="white", font=("Arial", 10))
        self.log_text.pack(fill="both", expand=True)

    def create_data_preview_table(self):
        """Create and place the data preview table."""
        self.table_frame = tk.Frame(self.main_frame, bg="#1e1e1e")
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def create_footer_buttons(self):
        """Create and place the footer buttons."""
        footer_frame = tk.Frame(self.main_frame, bg="#1e1e1e")
        footer_frame.pack(fill="x", pady=10)

        self.btn_ganti_tema = tk.Button(footer_frame, text="🎨 Ganti Tema", font=("Arial", 10, "bold"),
                                        bg="#0077B6", fg="white", padx=10, pady=5, relief="flat",
                                        command=self.ganti_tema)
        self.btn_ganti_tema.pack(side="left", padx=5)

        self.btn_keluar = tk.Button(footer_frame, text="❌ Keluar", font=("Arial", 10, "bold"),
                                    bg="#D9534F", fg="white", padx=10, pady=5, relief="flat",
                                    command=self.root.quit)
        self.btn_keluar.pack(side="right", padx=5)

    def log_message(self, message):
        """Log messages to the output area and console."""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        console.print(f"[bold green]{message}[/bold green]")

    def pilih_pdf(self):
        """Open file dialog to select PDF files."""
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if file_paths:
            for pdf_path in file_paths:
                self.ekstrak_pdf(pdf_path)

    def ekstrak_pdf(self, pdf_path):
        """Extract data from the selected PDF file."""
        self.log_message("🔄 Memulai ekstraksi data dari PDF...")

        try:
            output_dir = os.path.dirname(pdf_path)
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            all_data = []

            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        all_data.extend(table[1:])  # Skip header

            if not all_data:
                messagebox.showerror("Error", "❌ Tidak ada data yang diekstrak dari PDF.")
                self.log_message("❌ Tidak ada data ditemukan dalam PDF.")
                return

            header = ["No", "NIM", "Nama", "Kehadiran", "UTS", "UAS", "Sikap", "Quiz", "Tugas Kelompok", "Tugas Besar", "Nilai Akhir", "Mutu"]
            df = pd.DataFrame(all_data, columns=header)

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
        """Display extracted data in the console."""
        table = Table(title="📊 Hasil Ekstraksi PDF", show_lines=True)
        for col in df.columns:
            table.add_column(col, style="cyan")
        for _, row in df.iterrows():
            table.add_row(*[str(val) for val in row])
        console.print(table)

    def preview_data(self, df):
        """Display extracted data in a table within the GUI."""
        # Clear previous table if exists
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        # Create a PandasTable to display the DataFrame
        pt = PandasTable(self.table_frame, dataframe=df, showtoolbar=True, showstatusbar=True)
        pt.show()

    def ganti_tema(self):
        """Toggle between light and dark themes."""
        tema_baru = "darkly" if self.style.theme == "superhero" else "superhero"
        self.style.theme_use(tema_baru)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()