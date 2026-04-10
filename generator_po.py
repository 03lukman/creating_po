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
    # 1. Persiapan Path
    filename = f"PO_{so}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    final_path = os.path.join(save_dir, filename)
    logo_path = r"C:\Users\lukman\MAGANGHUB\po\gambar\logo_nashua.png"

    # Jalankan Excel secara invisible (latar belakang)
    app = xw.App(visible=False)
    
    try:
        # 2. Buka Template
        if template_path and os.path.exists(template_path):
            wb = app.books.open(template_path)
        else:
            print(f"❌ Template tidak ditemukan: {template_path}")
            return None
            
        sheet = wb.sheets[0]

        # 3. Pengisian Data
        if mode == "PROD":
            # Isi Header PO (E4 & E5)
            sheet.range("E4").value = generate_po_number(base_po, index)
            # Menggunakan tanggal dari database jika tersedia, jika tidak pakai hari ini
            tanggal_data = items[0][0] if items else datetime.now()
            sheet.range("E5").value = tanggal_data.strftime("%d/%m/%Y") if hasattr(tanggal_data, "strftime") else str(tanggal_data)

            # Isi Tabel Item (Mulai Baris 16)
            start_row = 16
            for i, item in enumerate(items):
                r = start_row + i
                
                # Definisikan cell utama
                cell_kode = sheet.range(f"A{r}")
                cell_desc = sheet.range(f"B{r}")
                cell_qty  = sheet.range(f"C{r}")
                cell_unit = sheet.range(f"D{r}")
                cell_price = sheet.range(f"E{r}")
                cell_amount = sheet.range(f"F{r}")

                # Tulis Nilai
                cell_kode.value = item[2]   # ITEMNO
                cell_desc.value = item[3]   # ITEMOVDESC (Deskripsi Panjang)
                cell_qty.value  = item[4]   # QUANTITY
                cell_unit.value = item[5]   # ITEMUNIT
                cell_price.value = item[-2] # HARGA
                cell_amount.value = item[-1] # TOTAL

                # --- STRATEGI PROFESIONAL: PENANGANAN DESKRIPSI PANJANG ---
                # Memastikan teks membungkus (wrap) dan baris melebar otomatis
                cell_desc.api.WrapText = True 
                
                # Mengatur perataan teks ke atas (Top) agar rapi jika deskripsi sangat panjang
                sheet.range(f"A{r}:F{r}").api.VerticalAlignment = xw.constants.VAlign.xlVAlignTop
                
                # Autofit hanya untuk baris yang sedang diisi
                sheet.range(f"{r}:{r}").rows.autofit()

                # Format Angka Rupiah/Ribuan
                cell_price.number_format = '#,##0'
                cell_amount.number_format = '#,##0'

        # 4. Finishing Logo (Header Sejati)
        if os.path.exists(logo_path):
            # Mengatur logo di sisi kiri (Left Header) sesuai identitas Nashua
            sheet.page_setup.left_header_picture = logo_path
            sheet.page_setup.left_header = '&G' 
            
            # Margin disesuaikan agar logo tidak menindih konten (Page Setup)
            sheet.page_setup.top_margin = 100   
            sheet.page_setup.header_margin = 20 
            
            # Optimasi cetak agar pas satu halaman lebar (A4)
            sheet.page_setup.zoom = False
            sheet.page_setup.fit_to_pages_wide = 1
            sheet.page_setup.fit_to_pages_tall = False # Biarkan memanjang ke bawah jika item sangat banyak
        else:
            print(f"⚠ Peringatan: Logo tidak ditemukan di {logo_path}")

        # 5. Simpan Hasil
        wb.save(final_path)
        wb.close()
        
    except Exception as e:
        print(f"❌ Error saat generate PO: {str(e)}")
        return None
    finally:
        # Pastikan aplikasi Excel benar-benar tertutup dari memori
        app.quit()

    return final_path