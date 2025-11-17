import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def cek_kolom_LM_satu_sheet():
    file_path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        excel_file = pd.ExcelFile(file_path)
        # Pilih nama sheet
        sheet_names = excel_file.sheet_names
        sheet = simpledialog.askstring(
            "Pilih Sheet", f"Masukkan nama sheet dari daftar ini:\n\n{', '.join(sheet_names)}"
        )

        if not sheet or sheet not in sheet_names:
            messagebox.showerror("Error", "Nama sheet tidak ditemukan di file.")
            return

        # Baca file Excel dengan header di baris ke-9 (index 8)
        df = pd.read_excel(file_path, sheet_name=sheet, header=8, usecols="A:P")

        # Pastikan kolom minimal 13 (karena L & M adalah kolom ke-12 dan ke-13)
        if df.shape[1] < 13:
            messagebox.showerror("Error", f"Sheet {sheet} tidak memiliki kolom L dan M (minimal 13 kolom).")
            return

        # Ambil kolom L dan M berdasarkan posisi
        kolom_L = df.columns[11]  # kolom ke-12
        kolom_M = df.columns[12]  # kolom ke-13

        # Filter baris dengan nilai L = M
        df_sama = df[df[kolom_L] == df[kolom_M]]

        if df_sama.empty:
            messagebox.showinfo("Info", f"Tidak ada baris dengan kolom L dan M sama di sheet {sheet}.")
        else:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Simpan Hasil Kolom L=M"
            )
            if save_path:
                df_sama.to_excel(save_path, index=False, sheet_name=sheet)
                messagebox.showinfo("Selesai", f"Hasil kolom L=M berhasil disimpan ke:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI
root = tk.Tk()
root.title("Cek Kolom L dan M Sama (1 Sheet)")
root.geometry("420x200")

tk.Label(root, text="Cek baris dengan nilai kolom L dan M yang sama\n(header di baris ke-9, kolom A–P)", 
         wraplength=380, justify="center").pack(pady=20)
tk.Button(root, text="Pilih File dan Jalankan", command=cek_kolom_LM_satu_sheet, 
          bg="#2196F3", fg="white", height=2).pack()
tk.Label(root, text="Hasil disimpan ke 1 file Excel (1 sheet)").pack(pady=10)

root.mainloop()
