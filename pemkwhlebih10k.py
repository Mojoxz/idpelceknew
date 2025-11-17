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

        # Sheet & header configuration
        sheet_headers = {
            "DMP": 7,
            "DKP": 6,
            "NGL": 6,
            "RKT": 6,
            "GDN": 6
        }

        hasil_semua_sheet = []

        for sheet, header_row in sheet_headers.items():
            try:
                df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)

                # Pastikan kolom PEMKWH ada
                if "PEMKWH" in df.columns:
                    filtered = df[df["PEMKWH"] > 10000]
                    filtered["SHEET"] = sheet  # Tandai asal sheet
                    hasil_semua_sheet.append(filtered)
            except Exception as e:
                print(f"Error membaca sheet {sheet}: {e}")

        # Gabungkan semua hasil
        if not hasil_semua_sheet:
            messagebox.showerror("Gagal", "Tidak ada data PEMKWH > 10000 ditemukan.")
            return

        final_df = pd.concat(hasil_semua_sheet, ignore_index=True)

        # Simpan hasil
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Simpan Hasil Filter"
        )

        if save_path:
            final_df.to_excel(save_path, index=False)
            messagebox.showinfo("Sukses", "Data berhasil disaring dan disimpan!")

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")


# GUI Tkinter
root = Tk()
root.title("Filter PEMKWH > 10000")
root.geometry("400x200")

Label(root, text="Filter Data PEMKWH > 10000", font=("Arial", 14, "bold")).pack(pady=20)
Button(root, text="Upload & Proses File Excel", font=("Arial", 12), command=proses_file).pack(pady=10)

root.mainloop()
