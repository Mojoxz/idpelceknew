import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def pilih_file_agst():
    global file_agst
    file_agst = filedialog.askopenfilename(
        title="Pilih File Excel Agustus",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_agst:
        lbl_agst.config(text=f"File Agustus: {file_agst}")

def pilih_file_sept():
    global file_sept
    file_sept = filedialog.askopenfilename(
        title="Pilih File Excel September",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_sept:
        lbl_sept.config(text=f"File September: {file_sept}")

def proses_data():
    try:
        if not file_agst or not file_sept:
            messagebox.showerror("Error", "Harap pilih kedua file Excel terlebih dahulu!")
            return

        # Baca kolom C dari kedua file
        df_agst = pd.read_excel(file_agst, usecols="C")
        df_sept = pd.read_excel(file_sept, usecols="C")

        # Normalisasi nama kolom
        df_agst.columns = ["IDPEL"]
        df_sept.columns = ["IDPEL"]

        # Cari IDPEL baru
        idpel_baru = df_sept[~df_sept["IDPEL"].isin(df_agst["IDPEL"])]

        # Minta user masukkan nama file
        nama_file = simpledialog.askstring("Simpan File", "Masukkan nama file hasil (tanpa .xlsx):")
        if not nama_file:
            messagebox.showwarning("Batal", "Penyimpanan dibatalkan.")
            return

        # Simpan ke Excel
        output_file = f"{nama_file}.xlsx"
        idpel_baru.to_excel(output_file, index=False)

        messagebox.showinfo("Selesai", f"Hasil IDPEL terbaru tersimpan di '{output_file}'")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI setup
root = tk.Tk()
root.title("Cek IDPEL Terbaru")
root.geometry("500x250")

lbl_agst = tk.Label(root, text="File Agustus: (belum dipilih)", wraplength=450)
lbl_agst.pack(pady=5)

btn_agst = tk.Button(root, text="Pilih File Agustus", command=pilih_file_agst)
btn_agst.pack(pady=5)

lbl_sept = tk.Label(root, text="File September: (belum dipilih)", wraplength=450)
lbl_sept.pack(pady=5)

btn_sept = tk.Button(root, text="Pilih File September", command=pilih_file_sept)
btn_sept.pack(pady=5)

btn_proses = tk.Button(root, text="Proses Data", command=proses_data, bg="green", fg="white")
btn_proses.pack(pady=20)

root.mainloop()
