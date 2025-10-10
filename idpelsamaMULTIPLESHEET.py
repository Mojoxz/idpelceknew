import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def cek_duplikat_idpel_multisheet():
    file_path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    target_sheets = ["DMP", "DKP", "NGL", "RKT", "GDN"]
    hasil_sheets = {}

    try:
        excel_file = pd.ExcelFile(file_path)

        for sheet in target_sheets:
            if sheet not in excel_file.sheet_names:
                continue

            df = pd.read_excel(file_path, sheet_name=sheet)

            # Cek apakah ada kolom bernama "IDPEL" (case-insensitive)
            kolom_idpel = [col for col in df.columns if str(col).strip().lower() == "idpel"]
            if not kolom_idpel:
                continue

            kolom_idpel = kolom_idpel[0]
            duplikat = df[df[kolom_idpel].duplicated(keep=False)]

            if not duplikat.empty:
                hasil_sheets[sheet] = duplikat

        if not hasil_sheets:
            messagebox.showinfo("Info", "Tidak ditemukan IDPEL duplikat di sheet yang ditentukan.")
            return

        # Simpan hasil ke file baru
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Simpan Hasil Duplikat Multi-sheet"
        )
        if save_path:
            with pd.ExcelWriter(save_path) as writer:
                for sheet, df_hasil in hasil_sheets.items():
                    df_hasil.to_excel(writer, index=False, sheet_name=sheet)
            messagebox.showinfo("Selesai", f"Hasil duplikat disimpan ke:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI
root = tk.Tk()
root.title("Cek Duplikat IDPEL (Multi-sheet)")
root.geometry("480x230")

tk.Label(
    root,
    text="Cek duplikat IDPEL pada beberapa sheet (DMP, DKP, NGL, RKT, GDN)\n"
         "Hasil per sheet disimpan terpisah di 1 file Excel",
    wraplength=450,
    justify="center"
).pack(pady=20)
tk.Button(
    root,
    text="Pilih File dan Jalankan",
    command=cek_duplikat_idpel_multisheet,
    bg="#4CAF50",
    fg="white",
    height=2
).pack()
tk.Label(root, text="Header akan dikenali otomatis dari kolom bernama 'IDPEL'").pack(pady=10)

root.mainloop()
