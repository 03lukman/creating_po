import xlwings as xw
import os
import re
from datetime import datetime

def generate_po_number(base, increment):
    """Menghasilkan nomor PO otomatis berdasarkan base PO dan index."""
    match = re.search(r'(\D+)(\d+)$', base)
    if not match: return base
    prefix, number = match.groups()
    # Menghitung nomor urut dan mempertahankan format 3 digit (misal: 023, 024)
    new_number = int(number) + increment
    return f"{prefix}{new_number:03d}"

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
            # --- 1. HEADER PO ---
            # Nomor PO dan Tanggal
            sheet.range("E4").value = generate_po_number(base_po, index)
            
            # Mengambil tanggal dari item pertama (hasil fetch DB)
            tanggal_data = items[0][0] if items else datetime.now()
            if hasattr(tanggal_data, "strftime"):
                sheet.range("E5").value = tanggal_data.strftime("%d/%m/%Y")
            else:
                sheet.range("E5").value = str(tanggal_data)

            # --- 2. ISI TABEL ITEM ---
            start_row = 16
            for i, item in enumerate(items):
                r = start_row + i
                
                # Mapping data hasil processor.py
                kode_item = item[2]
                deskripsi = str(item[3]).strip()
                qty       = item[4]
                unit      = item[5]
                # Index 6 adalah Ukuran murni hasil filter Excel
                ukuran    = str(item[6]).strip() if len(item) > 6 and item[6] else ""
                harga     = item[-2] # Harga DPP
                total     = item[-1] # Subtotal per item

                # Gabungkan Deskripsi dengan Ukuran (Newline)
                if ukuran and ukuran.lower() != "none":
                    full_text = f"{deskripsi}\nUkuran: {ukuran}"
                else:
                    full_text = deskripsi

                # Menulis data ke kolom Excel
                sheet.range(f"A{r}").value = kode_item
                sheet.range(f"B{r}").value = full_text
                sheet.range(f"C{r}").value = qty
                sheet.range(f"D{r}").value = unit
                sheet.range(f"E{r}").value = harga
                sheet.range(f"F{r}").value = total

                # --- 3. FORMATTING OTOMATIS ---
                target_cell_desc = sheet.range(f"B{r}")
                
                # Mengaktifkan Word Wrap agar teks panjang turun ke bawah (tidak lari ke kolom Ukuran)
                target_cell_desc.api.WrapText = True
                
                # Rata Atas agar kolom Kode/Qty sejajar dengan baris pertama Deskripsi
                sheet.range(f"A{r}:F{r}").api.VerticalAlignment = xw.constants.VAlign.xlVAlignTop
                
                # Auto-adjust tinggi baris setelah teks digabung
                sheet.range(f"{r}:{r}").rows.autofit()

                # Format angka ribuan tanpa desimal
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