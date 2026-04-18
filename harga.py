import openpyxl
import os

def find_header_row(sheet):
    # Cari baris yang mengandung "KODE BARANG" dan "NAMA BARANG"
    for i in range(1, 15):
        row_values = [str(cell.value).strip().upper() if cell.value else "" for cell in sheet[i]]
        if "KODE BARANG" in row_values and "NAMA BARANG" in row_values:
            return i
    raise Exception("Header tidak ditemukan di harga.xlsx")

def load_harga_dict(path):
    if not os.path.exists(path):
        raise Exception(f"File tidak ditemukan: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = wb.active 

    h_row = find_header_row(sheet)
    # Ambil header baris tersebut
    headers = [str(cell.value).strip().upper() if cell.value else "" for cell in sheet[h_row]]

    # Cari index kolom secara presisi
    try:
        idx_kode   = headers.index("KODE BARANG")
        idx_nama   = headers.index("NAMA BARANG")
        idx_ukuran = headers.index("UKURAN")
        idx_harga  = headers.index("HARGA DPP")
    except ValueError as e:
        raise Exception(f"Kolom tidak ditemukan: {e}")

    harga_dict = {}
    for row in sheet.iter_rows(min_row=h_row + 1, values_only=True):
        kode = str(row[idx_kode]).strip().upper() if row[idx_kode] else None
        
        if kode:
            # Ambil data murni dari kolom masing-masing
            nama_val = str(row[idx_nama]).strip() if row[idx_nama] else ""
            ukur_val = str(row[idx_ukuran]).strip() if row[idx_ukuran] else ""
            
            harga_dict[kode] = {
                "nama": nama_val,
                "ukuran": ukur_val,
                "harga": row[idx_harga] if row[idx_harga] is not None else 0
            }
    return harga_dict