import openpyxl
import os
import re
from datetime import datetime
from openpyxl.drawing.image import Image

# ===== HELPER: PO NUMBER =====
def generate_po_number(base, increment):
    match = re.search(r'(\D+)(\d+)$', base)
    if not match:
        return base

    prefix, number = match.groups()
    return f"{prefix}{int(number)+increment:03d}"


def generate_po(
    so,
    items,
    save_dir,
    mode="DEBUG",
    template_path=None,
    base_po="POR-NN26C023",
    index=0
):
    # ===== PILIH MODE TEMPLATE / NON TEMPLATE =====
    if template_path and mode == "PROD":
        wb = openpyxl.load_workbook(template_path)
        sheet = wb.active

    # ===== LOGO =====
        logo_path = r"C:\Users\lukman\MAGANGHUB\po\gambar\logo_nashua.png"

        if os.path.exists(logo_path):
            img = Image(logo_path)
            img.width = 250
            img.height = 80
            sheet.add_image(img, "A1")
        else:
            print(f"⚠ WARNING: Logo tidak ditemukan di {logo_path}")

    else:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "PO"

    # ===== VALIDASI SODATE =====
    dates = set(item[0] for item in items if item[0])
    if len(dates) > 1:
        print(f"⚠ WARNING: SO {so} punya multiple SODATE: {dates}")

    # ===== DATE =====
    tanggal = items[0][0]
    if hasattr(tanggal, "strftime"):
        tanggal_str = tanggal.strftime("%d/%m/%Y")
    else:
        tanggal_str = str(tanggal)

    # ===== MODE TEMPLATE =====
    if template_path and mode == "PROD":

        # PO NUMBER
        po_number = generate_po_number(base_po, index)
        sheet["E4"] = po_number

        # DATE
        sheet["E5"] = tanggal_str

        start_row = 16

        for i, item in enumerate(items):
            r = start_row + i

            sheet[f"A{r}"] = item[1]  # CONCAT
            sheet[f"B{r}"] = item[3]
            sheet[f"C{r}"] = item[4]
            sheet[f"D{r}"] = item[5]

            harga = item[-2]
            total = item[-1]

            if isinstance(harga, (int, float)):
                sheet[f"E{r}"] = harga
                sheet[f"E{r}"].number_format = '#,##0'
            else:
                sheet[f"E{r}"] = "NOT FOUND"

            if isinstance(total, (int, float)):
                sheet[f"F{r}"] = total
                sheet[f"F{r}"].number_format = '#,##0'

    # ===== MODE DEBUG (LAMA TETAP ADA) =====
    else:
        sheet["A1"] = "PURCHASE ORDER"
        sheet["A3"] = "NO SO:"
        sheet["B3"] = so

        sheet["A4"] = "DATE:"
        sheet["B4"] = tanggal_str

        start_row = 6

        sheet[f"A{start_row}"] = "KODE"
        sheet[f"B{start_row}"] = "NAMA"
        sheet[f"C{start_row}"] = "QTY"
        sheet[f"D{start_row}"] = "HARGA"
        sheet[f"E{start_row}"] = "TOTAL"

        for i, item in enumerate(items):
            r = start_row + 1 + i

            sheet[f"A{r}"] = item[2]
            sheet[f"B{r}"] = item[3]
            sheet[f"C{r}"] = item[4]
            sheet[f"C{r}"].number_format = '#,##0'

            harga = item[-2]
            total = item[-1]

            sheet[f"D{r}"] = harga if harga else "NOT FOUND"
            sheet[f"E{r}"] = total if total else ""

    # ===== SAVE =====
    filename = f"PO_{so}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    path = os.path.join(save_dir, filename)

    wb.save(path)
    return path