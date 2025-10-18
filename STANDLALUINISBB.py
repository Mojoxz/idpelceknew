import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

def pilih_file():
    file_path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

def sortir_data():
    file_path = entry_file.get()
    if not file_path or not os.path.exists(file_path):
        messagebox.showerror("Error", "File tidak ditemukan.")
        return
    
    try:
        # Ambil semua sheet
        xls = pd.ExcelFile(file_path)
        output_frames = []

        for sheet_name in xls.sheet_names:
            # Baca file dengan header di baris ke-8 (indeks 7)
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=7)

            # Cek apakah kolom SLALWBP dan SAHLWBP ada
            if "SLALWBP" in df.columns and "SAHLWBP" in df.columns:
                filtered = df[df["SLALWBP"] == df["SAHLWBP"]]
                if not filtered.empty:
                    filtered["SHEET"] = sheet_name  # tandai sheet asal
                    output_frames.append(filtered)

        if not output_frames:
            messagebox.showinfo("Hasil", "Tidak ada data dengan SLALWBP = SAHLWBP ditemukan.")
            return

        # Gabungkan hasil dari semua sheet
        hasil = pd.concat(output_frames, ignore_index=True)

        # Simpan ke file Excel baru
        output_path = os.path.join(os.path.dirname(file_path), "hasil_sortir_SLALWBP_SAHLWBP.xlsx")
        hasil.to_excel(output_path, index=False)

        messagebox.showinfo("Selesai", f"Hasil penyortiran disimpan di:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan:\n{e}")

# GUI Setup
root = tk.Tk()
root.title("Sortir SLALWBP = SAHLWBP - PLN Data")
root.geometry("600x200")
root.resizable(False, False)

frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

ttk.Label(frame, text="Pilih File Excel PLN:", font=("Arial", 11)).pack(anchor="w")

file_frame = ttk.Frame(frame)
file_frame.pack(fill="x", pady=5)

entry_file = ttk.Entry(file_frame, width=60)
entry_file.pack(side="left", padx=(0, 10), fill="x", expand=True)

ttk.Button(file_frame, text="Browse", command=pilih_file).pack(side="right")

ttk.Button(frame, text="Sortir Data (SLALWBP = SAHLWBP)", command=sortir_data).pack(pady=20)

root.mainloop()
