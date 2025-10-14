import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def process_file():
    file_path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        # Baca semua sheet target
        sheet_config = {
            "DMP": 7,   # header di baris ke-8 (index 7)
            "DKP": 6,   # header di baris ke-7 (index 6)
            "NGL": 6,
            "RKT": 6,
            "GDN": 6
        }

        output_writer = pd.ExcelWriter(f"{save_name_entry.get()}.xlsx", engine='openpyxl')

        for sheet, header_row in sheet_config.items():
            try:
                df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)
                if "SLALWBP" in df.columns and "SAHLWBP" in df.columns:
                    # Filter baris di mana SLALWBP == SAHLWBP
                    filtered = df[df["SLALWBP"] == df["SAHLWBP"]]
                    
                    if not filtered.empty:
                        filtered.to_excel(output_writer, sheet_name=sheet, index=False)
                    else:
                        # Jika tidak ada data yang cocok, tetap buat sheet kosong
                        empty_df = pd.DataFrame(columns=df.columns)
                        empty_df.to_excel(output_writer, sheet_name=sheet, index=False)
                else:
                    print(f"Kolom SLALWBP/SAHLWBP tidak ditemukan di sheet {sheet}")
            except Exception as e:
                print(f"Gagal memproses sheet {sheet}: {e}")

        output_writer.close()
        messagebox.showinfo("Sukses", f"Proses selesai!\nFile disimpan sebagai '{save_name_entry.get()}.xlsx'")
    
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")


# GUI Setup
root = tk.Tk()
root.title("Pengecek Data SLALWBP = SAHLWBP")
root.geometry("400x200")

label = tk.Label(root, text="Masukkan nama file hasil (tanpa .xlsx):", font=("Segoe UI", 10))
label.pack(pady=10)

save_name_entry = tk.Entry(root, width=40)
save_name_entry.pack(pady=5)
save_name_entry.insert(0, "Hasil_Pengecekan")

process_button = tk.Button(root, text="Pilih File Excel & Proses", command=process_file, bg="#4CAF50", fg="white", padx=10, pady=5)
process_button.pack(pady=20)

root.mainloop()
