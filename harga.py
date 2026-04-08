import openpyxl
import os


def find_header_row(sheet):
    for i in range(1, 10):
        row = [str(cell.value).upper() if cell.value else "" for cell in sheet[i]]
        if "KODE" in " ".join(row) and "HARGA" in " ".join(row):
            return i
    raise Exception("Header tidak ditemukan")


def load_harga_dict(path):
    if not os.path.exists(path):
        raise Exception(f"File harga tidak ditemukan: {path}")

    wb = openpyxl.load_workbook(path)
    sheet = wb.active

    header_row = find_header_row(sheet)
    headers = [str(cell.value).strip().upper() for cell in sheet[header_row]]

    def find_col(key):
        for i, h in enumerate(headers):
            if key in h:
                return i
        raise Exception(f"Kolom {key} tidak ditemukan")

    kode_idx = find_col("KODE")
    harga_idx = find_col("HARGA")

    harga_dict = {}

    for row in sheet.iter_rows(min_row=header_row+1, values_only=True):
        kode = str(row[kode_idx]).strip().upper() if row[kode_idx] else ""
        harga = row[harga_idx]

        if kode:
            harga_dict[kode] = harga

    return harga_dict