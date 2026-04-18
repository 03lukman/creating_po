def group_by_so(processed_data):
    """
    Mengelompokkan data berdasarkan Nomor SO (Index 1).
    Struktur item dalam processed_data:
    [0:Tanggal, 1:SO_Number, 2:Item_Code, 3:Nama_Barang, 4:Qty, 5:Unit, 6:Ukuran, 7:Harga, 8:Total]
    """
    grouped = {}
    
    for row in processed_data:
        # Pastikan kita mengambil SO Number dari index yang benar (Index 1)
        so_number = str(row[1]).strip()
        
        if so_number not in grouped:
            grouped[so_number] = []
        
        grouped[so_number].append(row)
        
    return grouped

def process_data(data, harga_dict):
    processed = []

    for row in data:
        row = list(row) 
        kode = str(row[2]).strip().upper() if row[2] else ""
        
        info = harga_dict.get(kode)
        
        if info:
            nama_barang = info["nama"]
            ukuran      = info["ukuran"]
            harga       = info["harga"]
        else:
            nama_barang = row[3]
            ukuran      = ""
            harga       = None

        qty = row[4]
        # Perhitungan total
        total = (float(harga) * float(qty)) if harga and qty else 0

        # MENGATUR ULANG INDEX AGAR KONSISTEN
        # Index 0-5 tetap dari DB (Tanggal, SO, Kode, Nama, Qty, Unit)
        row[3] = nama_barang   # Deskripsi murni
        
        # Gunakan indexing yang aman untuk menambahkan kolom baru
        # Kita ingin: [Tanggal(0), SO(1), Kode(2), Nama(3), Qty(4), Unit(5), Ukuran(6), Harga(7), Total(8)]
        
        if len(row) > 6:
            row[6] = ukuran
        else:
            row.append(ukuran) # Masuk ke index 6
            
        row.append(harga)      # Masuk ke index 7
        row.append(total)      # Masuk ke index 8

        processed.append(row)

    return processed