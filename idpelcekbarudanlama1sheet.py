import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

def baca_idpel(file_path):
    """Membaca semua sheet dan mengambil semua IDPEL dari kolom C (baris ke-9 ke bawah)."""
    try:
        excel_file = pd.ExcelFile(file_path)
        semua_idpel = set()
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet, header=None)
            # Pastikan ada minimal 9 baris dan 3 kolom
            if df.shape[0] > 9 and df.shape[1] > 2:
                # Cek apakah baris ke-9 (indeks 8) berisi "IDPEL"
                if str(df.iloc[8, 2]).strip().upper() == "IDPEL":
                    # Ambil kolom C mulai baris ke-10 (index 9)
                    kolom_idpel = df.iloc[9:, 2].dropna().astype(str).str.strip()
                    semua_idpel.update(kolom_idpel)
        return semua_idpel
    except Exception as e:
        messagebox.showerror("Error", f"Gagal membaca file {os.path.basename(file_path)}:\n{e}")
        return set()

def bandingkan():
    if not file_sep.get() or not file_okt.get():
        messagebox.showwarning("Peringatan", "Harap pilih kedua file terlebih dahulu.")
        return

    idpel_sep = baca_idpel(file_sep.get())
    idpel_okt = baca_idpel(file_okt.get())

    if not idpel_sep or not idpel_okt:
        messagebox.showwarning("Peringatan", "Tidak ditemukan data IDPEL pada salah satu file.")
        return

    baru_oktober = sorted(idpel_okt - idpel_sep)
    tidak_digunakan = sorted(idpel_sep - idpel_okt)

    hasil_text.delete("1.0", tk.END)
    hasil_text.insert(tk.END, f"📘 IDPEL BARU DI OKTOBER ({len(baru_oktober)}):\n")
    hasil_text.insert(tk.END, "\n".join(baru_oktober))
    hasil_text.insert(tk.END, "\n\n📙 IDPEL TIDAK DIGUNAKAN DI OKTOBER ({len(tidak_digunakan)}):\n")
    hasil_text.insert(tk.END, "\n".join(tidak_digunakan))

    global hasil_baru, hasil_tidak
    hasil_baru, hasil_tidak = baru_oktober, tidak_digunakan

def simpan_hasil():
    if not hasil_baru and not hasil_tidak:
        messagebox.showwarning("Peringatan", "Belum ada hasil untuk disimpan.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Simpan hasil perbandingan sebagai..."
    )
    if save_path:
        df_baru = pd.DataFrame(hasil_baru, columns=["IDPEL Baru di Oktober"])
        df_tidak = pd.DataFrame(hasil_tidak, columns=["IDPEL Tidak Digunakan di Oktober"])
        with pd.ExcelWriter(save_path) as writer:
            df_baru.to_excel(writer, index=False, sheet_name="IDPEL Baru")
            df_tidak.to_excel(writer, index=False, sheet_name="IDPEL Tidak Digunakan")
        messagebox.showinfo("Sukses", f"Hasil berhasil disimpan ke:\n{save_path}")

def pilih_file(var, label):
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if path:
        var.set(path)
        label.config(text=os.path.basename(path))

# ==== GUI ====
root = tk.Tk()
root.title("Perbandingan IDPEL Bulan September vs Oktober")
root.geometry("650x500")
root.resizable(False, False)

hasil_baru, hasil_tidak = [], []

frame = ttk.Frame(root, padding=10)
frame.pack(fill="both", expand=True)

# Pilih file September
file_sep = tk.StringVar()
file_okt = tk.StringVar()

ttk.Label(frame, text="Pilih File BULAN SEPTEMBER:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0,5))
lbl_sep = ttk.Label(frame, text="Belum dipilih", foreground="gray")
lbl_sep.pack(anchor="w")
ttk.Button(frame, text="Browse...", command=lambda: pilih_file(file_sep, lbl_sep)).pack(anchor="w", pady=(0,10))

# Pilih file Oktober
ttk.Label(frame, text="Pilih File BULAN OKTOBER:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0,5))
lbl_okt = ttk.Label(frame, text="Belum dipilih", foreground="gray")
lbl_okt.pack(anchor="w")
ttk.Button(frame, text="Browse...", command=lambda: pilih_file(file_okt, lbl_okt)).pack(anchor="w", pady=(0,10))

# Tombol bandingkan & simpan
ttk.Button(frame, text="🔍 Bandingkan Data", command=bandingkan).pack(pady=5)
ttk.Button(frame, text="💾 Simpan Hasil", command=simpan_hasil).pack(pady=5)

# Kotak hasil
hasil_text = tk.Text(frame, wrap="word", height=15)
hasil_text.pack(fill="both", expand=True, pady=10)

root.mainloop()
