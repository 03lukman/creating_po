def process_data(data, harga_dict):
    processed = []

    for row in data:
        row = list(row)

        kode = str(row[2]).strip().upper() if row[2] else ""
        harga = harga_dict.get(kode)

        qty = row[4]
        total = None

        try:
            if harga is not None and qty is not None:
                total = float(harga) * float(qty)
        except:
            total = None

        row.append(harga)
        row.append(total)

        processed.append(row)

    return processed

def group_by_so(data):
    grouped = {}

    for row in data:
        concat = str(row[1])  # CONCATENATION (SOA60013A)

        # ambil SO tanpa huruf terakhir
        so = concat[:-1]

        if so not in grouped:
            grouped[so] = []

        grouped[so].append(row)

    return grouped