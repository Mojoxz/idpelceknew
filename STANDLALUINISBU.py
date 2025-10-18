import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Variabel global untuk menyimpan hasil sortir terakhir
filtered_df_global = None

def load_excel():
    file_path = filedialog.askopenfilename(
        title="Pilih file Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_combo['values'] = excel_file.sheet_names
        sheet_combo.current(0)
        sheet_combo.file_path = file_path
        messagebox.showinfo("Sukses", f"File berhasil dimuat:\n{file_path}\n\nPilih sheet untuk diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal membuka file Excel:\n{e}")

def process_sheet():
    global filtered_df_global
    try:
        file_path = sheet_combo.file_path
        sheet_name = sheet_combo.get()
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)  # baris ke-9 jadi header

        if "STAND LALU" not in df.columns or "STAND INI" not in df.columns:
            messagebox.showerror("Error", "Kolom 'STAND LALU' atau 'STAND INI' tidak ditemukan.")
            return

        # Filter data yang STAND LALU ≠ STAND INI
        filtered_df = df[df["STAND LALU"] != df["STAND INI"]]
        filtered_df_global = filtered_df  # Simpan untuk tombol download

        # Tampilkan hasil ke Treeview
        for i in tree.get_children():
            tree.delete(i)

        tree["columns"] = list(filtered_df.columns)
        tree["show"] = "headings"

        for col in filtered_df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center")

        for _, row in filtered_df.iterrows():
            tree.insert("", "end", values=list(row))

        messagebox.showinfo("Selesai", f"Menampilkan {len(filtered_df)} data yang STAND LALU ≠ STAND INI")

    except Exception as e:
        messagebox.showerror("Error", f"Gagal memproses sheet:\n{e}")

def save_to_excel():
    global filtered_df_global
    if filtered_df_global is None or filtered_df_global.empty:
        messagebox.showwarning("Peringatan", "Belum ada data yang bisa disimpan.\nSilakan klik 'Periksa Sheet' dulu.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Simpan hasil sortir sebagai..."
    )
    if not file_path:
        return

    try:
        filtered_df_global.to_excel(file_path, index=False)
        messagebox.showinfo("Berhasil", f"Hasil sortir berhasil disimpan ke:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan file Excel:\n{e}")

# === GUI ===
root = tk.Tk()
root.title("Pemeriksa STAND LALU vs STAND INI (Excel)")
root.geometry("1100x650")

frame = tk.Frame(root)
frame.pack(pady=10)

load_btn = tk.Button(frame, text="📂 Pilih File Excel", command=load_excel)
load_btn.grid(row=0, column=0, padx=5)

sheet_combo = ttk.Combobox(frame, state="readonly", width=30)
sheet_combo.grid(row=0, column=1, padx=5)

process_btn = tk.Button(frame, text="🔍 Periksa Sheet", command=process_sheet)
process_btn.grid(row=0, column=2, padx=5)

save_btn = tk.Button(frame, text="💾 Simpan Hasil ke Excel", command=save_to_excel)
save_btn.grid(row=0, column=3, padx=5)

tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True)

tree_scroll = tk.Scrollbar(tree_frame)
tree_scroll.pack(side="right", fill="y")

tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set)
tree.pack(fill="both", expand=True)

tree_scroll.config(command=tree.yview)

root.mainloop()
