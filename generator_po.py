import xlwings as xw
import os
import re
from datetime import datetime

def generate_po_number(base, increment):
    match = re.search(r'(\D+)(\d+)$', base)
    if not match: return base
    prefix, number = match.groups()
    return f"{prefix}{int(number)+increment:03d}"

def generate_po(so, items, save_dir, mode="PROD", template_path=None, base_po="POR-NN26C023", index=0):
    # 1. Tentukan Path
    filename = f"PO_{so}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    final_path = os.path.join(save_dir, filename)
    logo_path = r"C:\Users\lukman\MAGANGHUB\po\gambar\logo_nashua.png"

    # Jalankan Excel di latar belakang
    app = xw.App(visible=False)
    
    try:
        # 2. Buka Template Langsung
        if template_path and os.path.exists(template_path):
            wb = app.books.open(template_path)
        else:
            print("❌ Template tidak ditemukan!")
            return None
            
        sheet = wb.sheets[0]

        # 3. ISI DATA KE SEL (Sama seperti cara openpyxl tapi versi xlwings)
        if mode == "PROD":
            # Isi Header PO
            sheet.range("E4").value = generate_po_number(base_po, index)
            sheet.range("E5").value = items[0][0].strftime("%d/%m/%Y") if hasattr(items[0][0], "strftime") else str(items[0][0])

            # Isi Tabel Item (Mulai baris 16)
            start_row = 16
            for i, item in enumerate(items):
                r = start_row + i
                sheet.range(f"A{r}").value = item[1] # Kode
                sheet.range(f"B{r}").value = item[3] # Description
                sheet.range(f"C{r}").value = item[4] # Qty
                sheet.range(f"D{r}").value = item[5] # Unit
                sheet.range(f"E{r}").value = item[-2] # Price
                sheet.range(f"F{r}").value = item[-1] # Amount
                
                # Format Angka (Ribuan pakai koma/titik)
                sheet.range(f"E{r}:F{r}").number_format = '#,##0'

        # 4. SETTING HEADER GAMBAR (FINISHING)
        if os.path.exists(logo_path):
            # Penting: Masukkan ke Left Header karena posisi Nashua di kiri
            sheet.page_setup.left_header_picture = logo_path
            sheet.page_setup.left_header = '&G' 
            
            # Atur Margin agar logo punya ruang dan tidak menabrak tabel
            sheet.page_setup.top_margin = 100   # Jarak konten dari atas kertas
            sheet.page_setup.header_margin = 20 # Jarak header dari atas kertas
            
            # Paksa agar pas di satu halaman A4
            sheet.page_setup.zoom = False
            sheet.page_setup.fit_to_pages_wide = 1
            sheet.page_setup.fit_to_pages_tall = 1
        else:
            print(f"⚠ Logo tidak ditemukan di: {logo_path}")

        # 5. SIMPAN SEBAGAI FILE BARU (Agar template asli tidak berubah)
        wb.save(final_path)
        wb.close()
        
    except Exception as e:
        print(f"❌ Terjadi kesalahan: {e}")
        return None
    finally:
        # Pastikan Excel tertutup sempurna
        app.quit()

    return final_path