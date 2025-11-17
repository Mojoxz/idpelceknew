# ----------------------------------------------------------------------------
# PASTIKAN ANDA MENGGUNAKAN VERSI TERBARU LIBRARY pdf2docx
# Jalankan perintah ini di terminal/CMD Anda: pip install --upgrade pdf2docx
# ----------------------------------------------------------------------------

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess # Tambahkan import ini untuk membuka folder
from pdf2docx import Converter

# --- Kelas Utama Aplikasi ---
class PdfConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Konverter PDF ke Word")
        self.root.geometry("600x250")
        self.root.resizable(False, False)

        # Variabel untuk menyimpan path file
        self.pdf_path = tk.StringVar()

        # Membuat Frame utama
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Elemen-elemen GUI ---
        
        # Label Judul
        title_label = ttk.Label(self.main_frame, text="Konverter PDF ke Word (.docx)", font=("Helvetica", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Label instruksi
        instruction_label = ttk.Label(self.main_frame, text="Pilih file PDF yang akan dikonversi:")
        instruction_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)

        # Entry untuk menampilkan path file
        self.path_entry = ttk.Entry(self.main_frame, textvariable=self.pdf_path, width=50, state="readonly")
        self.path_entry.grid(row=2, column=0, columnspan=2, padx=(0, 10), pady=5, sticky="ew")

        # Tombol Browse
        browse_button = ttk.Button(self.main_frame, text="Browse...", command=self.browse_file)
        browse_button.grid(row=2, column=2, pady=5, sticky="e")

        # Tombol Konversi
        self.convert_button = ttk.Button(self.main_frame, text="Konversi ke Word", command=self.convert_file, state="disabled")
        self.convert_button.grid(row=3, column=0, columnspan=3, pady=20)

        # Label Status
        self.status_label = ttk.Label(self.main_frame, text="Silakan pilih file PDF.", font=("Helvetica", 10))
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)

        # Progress Bar (indeterminate)
        self.progress = ttk.Progressbar(self.main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=5)

        # Mengatur konfigurasi grid agar entry bisa melebar
        self.main_frame.columnconfigure(0, weight=1)

    def browse_file(self):
        """Membuka dialog untuk memilih file PDF."""
        # Memfilter hanya file PDF
        filename = filedialog.askopenfilename(
            title="Pilih file PDF",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if filename:
            self.pdf_path.set(filename)
            self.status_label.config(text="File PDF dipilih. Klik 'Konversi ke Word' untuk memulai.")
            self.convert_button.config(state="normal")

    def convert_file(self):
        """Melakukan proses konversi dari PDF ke DOCX."""
        input_pdf = self.pdf_path.get()
        if not input_pdf:
            messagebox.showerror("Error", "File PDF belum dipilih!")
            return

        # Menentukan path output (nama file sama, ekstensi .docx)
        output_docx = os.path.splitext(input_pdf)[0] + '.docx'

        # Update UI status
        self.status_label.config(text="Sedang mengonversi... Mohon tunggu.")
        self.convert_button.config(state="disabled")
        self.progress.start(10) # Memulai animasi progress bar

        try:
            # Membuat objek konverter
            cv = Converter(input_pdf)
            # Melakukan konversi
            cv.convert(output_docx, start=0, end=None) # end=None untuk mengonversi semua halaman
            # Menutup konverter untuk melepas file
            cv.close()

            self.progress.stop() # Menghentikan progress bar
            self.status_label.config(text=f"Konversi berhasil! File tersimpan di:\n{output_docx}")
            
            # Menampilkan dialog konfirmasi sukses
            result = messagebox.askyesno(
                "Konversi Berhasil", 
                f"File berhasil dikonversi dan disimpan sebagai:\n{os.path.basename(output_docx)}\n\nApakah Anda ingin membuka folder penyimpanan?"
            )
            
            if result:
                # Membuka folder tempat file disimpan (cross-platform)
                try:
                    if os.name == 'nt': # Untuk Windows
                        os.startfile(os.path.dirname(output_docx))
                    elif os.name == 'posix': # Untuk macOS dan Linux
                        subprocess.run(['open', os.path.dirname(output_docx)])
                except Exception as e:
                    messagebox.showwarning("Peringatan", f"Tidak dapat membuka folder secara otomatis.\nError: {e}")

        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Terjadi kesalahan selama konversi.")
            messagebox.showerror("Kesalahan Konversi", f"Gagal mengonversi file.\n\nPesan Error: {e}")
        finally:
            # Mengembalikan tombol konversi ke keadaan normal
            self.convert_button.config(state="normal")


# --- Menjalankan Aplikasi ---
if __name__ == "__main__":
    root = tk.Tk()
    app = PdfConverterApp(root)
    root.mainloop()