import openpyxl
import os
from datetime import datetime


def generate_po(so, items, save_dir, mode="DEBUG"):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "PO"

    # ===== VALIDASI SODATE =====
    dates = set()
    for item in items:
        if item[0]:
            dates.add(item[0])

    if len(dates) > 1:
        if mode == "DEBUG":
            print(f"⚠ WARNING: SO {so} punya multiple SODATE: {dates}")
        else:
            raise Exception(f"❌ ERROR: SO {so} memiliki lebih dari 1 SODATE: {dates}")

    # ===== HEADER =====
    sheet["A1"] = "PURCHASE ORDER"

    sheet["A3"] = "NO SO:"
    sheet["B3"] = so

    sheet["A4"] = "DATE:"

    tanggal = items[0][0]
    if hasattr(tanggal, "strftime"):
        sheet["B4"] = tanggal.strftime("%d/%m/%Y")
    else:
        sheet["B4"] = str(tanggal)

    # ===== TABLE HEADER =====
    start_row = 6

    sheet[f"A{start_row}"] = "KODE"
    sheet[f"B{start_row}"] = "NAMA"
    sheet[f"C{start_row}"] = "QTY"
    sheet[f"D{start_row}"] = "HARGA"
    sheet[f"E{start_row}"] = "TOTAL"

    # ===== SORT ITEM =====
    items = sorted(items, key=lambda x: x[2])

    # ===== DATA =====
    for i, item in enumerate(items):
        r = start_row + 1 + i

        sheet[f"A{r}"] = item[2]   # ITEMNO
        sheet[f"B{r}"] = item[3]   # DESC
        sheet[f"C{r}"] = item[4]   # QTY
        sheet[f"C{r}"].number_format = '#,##0'

        harga = item[-2]
        total = item[-1]

        # ===== HARGA =====
        if isinstance(harga, (int, float)):
            sheet[f"D{r}"] = harga
            sheet[f"D{r}"].number_format = '#,##0'
        else:
            sheet[f"D{r}"] = "NOT FOUND"

        # ===== TOTAL =====
        if isinstance(total, (int, float)):
            sheet[f"E{r}"] = total
            sheet[f"E{r}"].number_format = '#,##0'
        else:
            sheet[f"E{r}"] = ""

    # ===== SUBTOTAL =====
    subtotal = 0
    for item in items:
        if isinstance(item[-1], (int, float)):
            subtotal += item[-1]

    last_row = start_row + len(items) + 2

    sheet[f"D{last_row}"] = "TOTAL"
    sheet[f"E{last_row}"] = subtotal
    sheet[f"E{last_row}"].number_format = '#,##0'

    # ===== SAVE =====
    filename = f"PO_{so}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    path = os.path.join(save_dir, filename)

    wb.save(path)

    return path