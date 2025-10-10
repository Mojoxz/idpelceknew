import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def cari_idpel_duplikat():
    # Pilih file Excel
    file_path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not file_path:
        return
    
    try:
        excel_file = pd.ExcelFile(file_path)
        semua_data = []

        for sheet in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet, header=None)
            
            # Cari kolom yang berisi header "IDPEL"
            idpel_col = None
            for i in range(len(df.columns)):
                if df.iloc[:, i].astype(str).str.contains("IDPEL", case=False, na=False).any():
                    idpel_col = i
                    break
            
            if idpel_col is not None:
                # Ambil data mulai setelah baris header IDPEL
                start_row = df[df.iloc[:, idpel_col].astype(str).str.contains("IDPEL", case=False, na=False)].index[0]
                data_idpel = df.iloc[start_row+1:, idpel_col].dropna().astype(str)
                temp_df = pd.DataFrame({"IDPEL": data_idpel, "Sheet": sheet})
                semua_data.append(temp_df)

        if not semua_data:
            messagebox.showerror("Error", "Tidak ditemukan kolom IDPEL di file ini.")
            return
        
        gabung = pd.concat(semua_data, ignore_index=True)

        # Cari IDPEL yang duplikat
        duplikat = gabung[gabung.duplicated(subset=["IDPEL"], keep=False)].sort_values(by="IDPEL")

        if duplikat.empty:
            messagebox.showinfo("Hasil", "Tidak ada IDPEL yang duplikat.")
        else:
            # Tampilkan hasil di tabel GUI
            for row in tree.get_children():
                tree.delete(row)
            for _, row in duplikat.iterrows():
                tree.insert("", "end", values=(row["IDPEL"], row["Sheet"]))

            simpan = messagebox.askyesno("Simpan Hasil", "Apakah ingin menyimpan hasil duplikat ke file Excel?")
            if simpan:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Simpan Hasil Sebagai"
                )
                if save_path:
                    duplikat.to_excel(save_path, index=False)
                    messagebox.showinfo("Berhasil", f"Hasil duplikat disimpan ke:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Gagal memproses file:\n{e}")

# === GUI ===
root = tk.Tk()
root.title("Cek IDPEL Duplikat PLN")
root.geometry("600x400")
root.resizable(False, False)

frame = tk.Frame(root)
frame.pack(pady=10)

btn_pilih = tk.Button(frame, text="Pilih File Excel dan Cek Duplikat", command=cari_idpel_duplikat, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
btn_pilih.pack(pady=5)

# Tabel hasil
columns = ("IDPEL", "Sheet")
tree = ttk.Treeview(root, columns=columns, show="headings", height=12)
tree.heading("IDPEL", text="IDPEL")
tree.heading("Sheet", text="Sheet")
tree.column("IDPEL", width=200)
tree.column("Sheet", width=100)
tree.pack(pady=10, fill="x")

root.mainloop()
  