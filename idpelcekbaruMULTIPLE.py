import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def pilih_file(label):
    path = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if path:
        label.config(text=path)
    return path

def bandingkan_data():
    file_okt = label_okt.cget("text")
    file_sep = label_sep.cget("text")

    if not file_okt or not file_sep or file_okt == "Belum dipilih" or file_sep == "Belum dipilih":
        messagebox.showwarning("Peringatan", "Pilih kedua file terlebih dahulu!")
        return

    sheets = ["DMP", "RKT", "GDN", "NGL", "DKP"]
    hasil = {}

    try:
        for sheet in sheets:
            okt = pd.read_excel(file_okt, sheet_name=sheet, header=None)
            sep = pd.read_excel(file_sep, sheet_name=sheet, header=None)

            # Ambil kolom C mulai dari baris yang sesuai
            start_row = 8 if sheet == "DMP" else 7
            idpel_okt = okt.iloc[start_row:, 2].dropna().astype(str).unique()
            idpel_sep = sep.iloc[start_row:, 2].dropna().astype(str).unique()

            # Deteksi IDPEL baru dan tidak digunakan
            idpel_baru = sorted(set(idpel_okt) - set(idpel_sep))
            idpel_hilang = sorted(set(idpel_sep) - set(idpel_okt))

            hasil[sheet] = {
                "baru": idpel_baru,
                "hilang": idpel_hilang
            }

        # Kosongkan tabel GUI
        for row in tree.get_children():
            tree.delete(row)

        # Tampilkan di GUI
        for sheet, data in hasil.items():
            max_len = max(len(data["baru"]), len(data["hilang"]))
            if max_len == 0:
                tree.insert("", "end", values=(sheet, "-", "Tidak ada perubahan"))
            else:
                for i in range(max_len):
                    val_baru = data["baru"][i] if i < len(data["baru"]) else ""
                    val_hilang = data["hilang"][i] if i < len(data["hilang"]) else ""
                    tree.insert("", "end", values=(sheet, val_baru, val_hilang))

        messagebox.showinfo("Selesai", "Perbandingan selesai! Anda dapat menyimpan hasilnya.")
        global hasil_per_sheet
        hasil_per_sheet = hasil

    except Exception as e:
        messagebox.showerror("Error", str(e))


def simpan_hasil():
    if 'hasil_per_sheet' not in globals():
        messagebox.showwarning("Peringatan", "Lakukan perbandingan terlebih dahulu!")
        return

    save_path = filedialog.asksaveasfilename(
        title="Simpan Hasil Sebagai",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not save_path:
        return

    try:
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            for sheet, data in hasil_per_sheet.items():
                df_baru = pd.DataFrame({"IDPEL Baru (Oktober)": data["baru"]})
                df_hilang = pd.DataFrame({"IDPEL Tidak Digunakan (Hilang di Oktober)": data["hilang"]})

                # Gabungkan dua kolom menjadi satu sheet
                max_len = max(len(df_baru), len(df_hilang))
                df_baru = df_baru.reindex(range(max_len))
                df_hilang = df_hilang.reindex(range(max_len))
                df_out = pd.concat([df_baru, df_hilang], axis=1)

                df_out.to_excel(writer, index=False, sheet_name=sheet)

        messagebox.showinfo("Berhasil", f"Hasil perbandingan telah disimpan di:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# ==== GUI ====
root = tk.Tk()
root.title("Perbandingan IDPEL September vs Oktober")
root.geometry("900x550")

frm = tk.Frame(root, padx=10, pady=10)
frm.pack(fill="x")

tk.Label(frm, text="File Oktober:").grid(row=0, column=0, sticky="w")
label_okt = tk.Label(frm, text="Belum dipilih", fg="gray")
label_okt.grid(row=0, column=1, sticky="w", padx=5)
tk.Button(frm, text="Pilih File", command=lambda: pilih_file(label_okt)).grid(row=0, column=2, padx=5)

tk.Label(frm, text="File September:").grid(row=1, column=0, sticky="w")
label_sep = tk.Label(frm, text="Belum dipilih", fg="gray")
label_sep.grid(row=1, column=1, sticky="w", padx=5)
tk.Button(frm, text="Pilih File", command=lambda: pilih_file(label_sep)).grid(row=1, column=2, padx=5)

tk.Button(frm, text="Bandingkan Data", command=bandingkan_data, bg="#2196F3", fg="white").grid(row=2, column=1, pady=10)
tk.Button(frm, text="Simpan Hasil", command=simpan_hasil, bg="#4CAF50", fg="white").grid(row=2, column=2, pady=10)

# Tabel hasil
tree = ttk.Treeview(root, columns=("Sheet", "IDPEL Baru", "IDPEL Tidak Digunakan"), show="headings")
tree.heading("Sheet", text="Sheet")
tree.heading("IDPEL Baru", text="IDPEL Baru (Oktober)")
tree.heading("IDPEL Tidak Digunakan", text="IDPEL Tidak Digunakan (Hilang di Oktober)")
tree.pack(fill="both", expand=True, padx=10, pady=10)

root.mainloop()
