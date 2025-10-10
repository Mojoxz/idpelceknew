import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def cek_duplikat_idpel_satu_sheet():
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

        df = pd.read_excel(file_path, sheet_name=sheet)
        if 'IDPEL' not in df.columns:
            messagebox.showerror("Error", f"Tidak ada kolom 'IDPEL' di sheet {sheet}.")
            return

        duplikat = df[df.duplicated(subset='IDPEL', keep=False)]

        if duplikat.empty:
            messagebox.showinfo("Info", f"Tidak ditemukan IDPEL duplikat di sheet {sheet}.")
        else:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Simpan Hasil Duplikat"
            )
            if save_path:
                duplikat.to_excel(save_path, index=False, sheet_name=sheet)
                messagebox.showinfo("Selesai", f"Hasil duplikat berhasil disimpan ke:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI
root = tk.Tk()
root.title("Cek Duplikat IDPEL (1 Sheet)")
root.geometry("420x200")

tk.Label(root, text="Cek Duplikat IDPEL (Kolom C)\nuntuk 1 sheet yang dipilih", wraplength=380, justify="center").pack(pady=20)
tk.Button(root, text="Pilih File dan Jalankan", command=cek_duplikat_idpel_satu_sheet, bg="#4CAF50", fg="white", height=2).pack()
tk.Label(root, text="Hasil disimpan ke 1 file Excel (1 sheet)").pack(pady=10)

root.mainloop()
