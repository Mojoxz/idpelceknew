import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def cek_kolom_LM():
    file_path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    sheets_to_check = ["DMP", "DKP", "NGL", "RKT", "GDN"]
    hasil_sheets = {}

    try:
        excel_file = pd.ExcelFile(file_path)
        for sheet in sheets_to_check:
            if sheet in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)

                # Pastikan ada minimal 13 kolom (sampai M)
                if df.shape[1] >= 13:
                    kolom_L = df.columns[11]
                    kolom_M = df.columns[12]
                    df_sama = df[df[kolom_L] == df[kolom_M]]

                    if not df_sama.empty:
                        hasil_sheets[sheet] = df_sama

        if hasil_sheets:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Simpan Hasil Kolom L=M (Multi Sheet)"
            )
            if save_path:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    for sheet_name, data in hasil_sheets.items():
                        data.to_excel(writer, index=False, sheet_name=sheet_name)
                messagebox.showinfo("Selesai", f"Hasil kolom L=M berhasil disimpan ke:\n{save_path}")
        else:
            messagebox.showinfo("Info", "Tidak ada baris dengan kolom L dan M yang sama di sheet yang diperiksa.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI
root = tk.Tk()
root.title("Cek Kolom L dan M Sama (Per Sheet)")
root.geometry("420x200")

tk.Label(root, text="Cek baris dengan nilai kolom L dan M yang sama\nHasil tetap per sheet: DMP, DKP, NGL, RKT, GDN", wraplength=380, justify="center").pack(pady=20)
tk.Button(root, text="Pilih File Excel dan Jalankan", command=cek_kolom_LM, bg="#2196F3", fg="white", height=2).pack()
tk.Label(root, text="Hasil disimpan dalam satu file, tapi tiap sheet terpisah").pack(pady=10)

root.mainloop()
