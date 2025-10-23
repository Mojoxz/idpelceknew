import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def cek_sama_kolom_IJ_satu_sheet():
    file_path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        excel_file = pd.ExcelFile(file_path)
        # Tampilkan daftar sheet untuk dipilih
        sheet_names = excel_file.sheet_names
        sheet = simpledialog.askstring("Pilih Sheet", f"Masukkan nama sheet dari daftar ini:\n\n{', '.join(sheet_names)}")

        if not sheet or sheet not in sheet_names:
            messagebox.showerror("Error", "Nama sheet tidak ditemukan di file.")
            return

        # Baca sheet tanpa header (data langsung)
        df = pd.read_excel(file_path, sheet_name=sheet, header=None)

        # Pastikan kolom I dan J (kolom ke 9 dan 10) ada
        if df.shape[1] < 10:
            messagebox.showerror("Error", f"Sheet {sheet} tidak memiliki kolom I dan J.")
            return

        # Bandingkan kolom I dan J
        kondisi_sama = df.iloc[:, 8] == df.iloc[:, 9]

        # Filter baris yang kolom I dan J-nya sama
        hasil_sama = df[kondisi_sama]

        if hasil_sama.empty:
            messagebox.showinfo("Info", f"Tidak ditemukan data yang sama antara kolom I dan J di sheet {sheet}.")
        else:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Simpan Hasil (Baris Sama I=J)"
            )
            if save_path:
                hasil_sama.to_excel(save_path, index=False, header=False, sheet_name=sheet)
                messagebox.showinfo("Selesai", f"Hasil baris yang sama disimpan ke:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI
root = tk.Tk()
root.title("Cek Data Sama Kolom I dan J (1 Sheet)")
root.geometry("460x220")

tk.Label(
    root,
    text="Cek baris di mana kolom I dan J memiliki nilai yang sama\n"
         "Menampilkan semua kolom dari A sampai Q",
    wraplength=400,
    justify="center"
).pack(pady=20)
tk.Button(
    root,
    text="Pilih File dan Jalankan",
    command=cek_sama_kolom_IJ_satu_sheet,
    bg="#2196F3",
    fg="white",
    height=2
).pack()
tk.Label(root, text="Hasil disimpan ke file Excel baru (1 sheet)").pack(pady=10)

root.mainloop()
