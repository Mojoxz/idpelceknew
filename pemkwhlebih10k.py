import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox

def proses_file():
    try:
        file_path = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        # Konfigurasi sheet & baris header (0-based index)
        sheet_headers = {
            "DMP": 7,
            "DKP": 6,
            "NGL": 6,
            "RKT": 6,
            "GDN": 6
        }

        # Pilih lokasi simpan
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Simpan Hasil Filter"
        )
        if not save_path:
            return

        # Open writer
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:

            any_written = False

            for sheet, header_row in sheet_headers.items():
                try:
                    df = pd.read_excel(
                        file_path, 
                        sheet_name=sheet,
                        header=header_row,
                        usecols="A:Q"   # ⬅️ hanya ambil kolom A sampai Q
                    )

                    if "PEMKWH" not in df.columns:
                        print(f"Sheet '{sheet}' tidak punya kolom PEMKWH. Dilewati.")
                        continue

                    # Konversi ke angka
                    df["PEMKWH"] = pd.to_numeric(df["PEMKWH"], errors="coerce")

                    # Filter
                    filtered = df[df["PEMKWH"] > 10000].copy()

                    if not filtered.empty:
                        filtered.loc[:, "SHEET"] = sheet

                        # Nama sheet tidak boleh lebih dari 31 karakter
                        sheetname_out = f"{sheet}_filtered"
                        if len(sheetname_out) > 31:
                            sheetname_out = sheetname_out[:31]

                        filtered.to_excel(writer, index=False, sheet_name=sheetname_out)
                        any_written = True

                except Exception as e:
                    print(f"Error membaca sheet {sheet}: {e}")

        if any_written:
            messagebox.showinfo("Sukses", f"Hasil berhasil disimpan di:\n{save_path}")
        else:
            messagebox.showinfo("Kosong", "Tidak ada data PEMKWH > 10000 ditemukan.")

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")


# GUI Tkinter
root = Tk()
root.title("Filter PEMKWH > 10000")
root.geometry("420x200")

Label(root, text="Filter Data PEMKWH > 10000 (Kolom A–Q)", font=("Arial", 14, "bold")).pack(pady=18)
Button(root, text="Upload & Proses File Excel", font=("Arial", 12), command=proses_file).pack(pady=10)
Label(root, text="Output: 1 file, banyak sheet hasil filter", wraplength=380, justify="center").pack(pady=8)

root.mainloop()
