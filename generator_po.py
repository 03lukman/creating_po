import xlwings as xw
import os
import re
import textwrap
from datetime import datetime

def generate_po_number(base, increment):
    """Logika yang sudah work: Menghasilkan nomor PO otomatis."""
    match = re.search(r'(\D+)(\d+)$', base)
    if not match: return base
    prefix, number = match.groups()
    new_number = int(number) + increment
    return f"{prefix}{new_number:03d}"

def wrap_kode(text, width=9):
    """Logika khusus Kode: Potong paksa setiap 10 karakter (tanpa cari spasi)."""
    if not text: return ""
    text = str(text)
    return "\n".join([text[i:i+width] for i in range(0, len(text), width)])

def wrap_deskripsi(text, width=35):
    """Logika khusus Deskripsi: Potong rapi per 35 karakter (mencari spasi)."""
    if not text: return ""
    lines = textwrap.wrap(str(text), width=width, break_long_words=True)
    return "\n".join(lines)

def generate_po(so, items, save_dir, mode="PROD", template_path=None, base_po="POR-NN26C023", index=0):
    filename = f"PO_{so}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    final_path = os.path.join(save_dir, filename)
    app = xw.App(visible=False)
    
    try:
        if template_path and os.path.exists(template_path):
            wb = app.books.open(template_path)
        else:
            print(f"❌ Template tidak ditemukan: {template_path}")
            return None
            
        sheet = wb.sheets[0]

        if mode == "PROD":
            # --- 1. HEADER PO (LOGIKA WORK) ---
            sheet.range("E4").value = generate_po_number(base_po, index)
            
            tanggal_data = items[0][0] if items else datetime.now()
            if hasattr(tanggal_data, "strftime"):
                sheet.range("E5").value = tanggal_data.strftime("%d/%m/%Y")
            else:
                sheet.range("E5").value = str(tanggal_data)

            # --- 2. ISI TABEL ITEM ---
            start_row = 16
            for i, item in enumerate(items):
                r = start_row + i
                
                # Gunakan logika berbeda untuk Kode vs Deskripsi
                kode_item = wrap_kode(item[2], width=9)
                
                deskripsi_raw = str(item[3]).strip()
                deskripsi_wrapped = wrap_deskripsi(deskripsi_raw, width=35)
                
                qty    = item[4]
                unit   = item[5]
                ukuran = str(item[6]).strip() if len(item) > 6 and item[6] else ""
                harga  = item[-2]
                total  = item[-1]

                # Gabungkan deskripsi rapi dengan ukuran
                if ukuran and ukuran.lower() != "none":
                    full_text = f"{deskripsi_wrapped}\nUkuran: {ukuran}"
                else:
                    full_text = deskripsi_wrapped

                # Menulis data ke kolom Excel
                sheet.range(f"A{r}").value = kode_item
                sheet.range(f"B{r}").value = full_text
                sheet.range(f"C{r}").value = qty
                sheet.range(f"D{r}").value = unit
                sheet.range(f"E{r}").value = harga
                sheet.range(f"F{r}").value = total

                # --- 3. FIXING LAYOUT & STABILITAS ---
                # Kunci tinggi baris agar volume tetap (Fixed Height)
                sheet.range(f"{r}:{r}").row_height = 40 
                
                target_range = sheet.range(f"A{r}:F{r}")
                target_range.api.WrapText = True
                target_range.api.VerticalAlignment = xw.constants.VAlign.xlVAlignTop
                
                # Format angka ribuan
                sheet.range(f"E{r}:F{r}").number_format = '#,##0'

        # --- 4. PENYIMPANAN ---
        wb.save(final_path)
        wb.close()
        
    except Exception as e:
        print(f"❌ Error Detail pada Generator: {str(e)}")
        return None
    finally:
        app.quit()

    return final_path