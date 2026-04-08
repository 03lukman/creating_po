import os
from dotenv import load_dotenv

from db import fetch_data
from harga import load_harga_dict
from processor import process_data, group_by_so
from generator_po import generate_po
from exporter import export_excel

MODE = os.getenv("MODE", "DEBUG")  # DEBUG / PROD
# ===== LOAD ENV =====
load_dotenv()

SAVE_DIR = os.getenv('SAVE_DIR', r"C:\Users\lukman\MAGANGHUB\po\test")
HARGA_PATH = r"C:\Users\lukman\MAGANGHUB\po\harga.xlsx"


def main():
    print("🔄 Load harga...")
    harga_dict = load_harga_dict(HARGA_PATH)

    print("🔄 Fetch data...")
    data, columns = fetch_data()

    if not data:
        print("❌ Tidak ada data")
        return

    print("🔄 Process data...")
    processed = process_data(data, harga_dict)

    print("🔄 Grouping per SO...")
    grouped = group_by_so(processed)

    print(f"✅ Total SO: {len(grouped)}")

    # ===== VALIDASI GROUPING =====
    print("\n===== VALIDASI GROUPING =====")

    total_items = 0

    for so, items in grouped.items():
        print(f"\nSO: {so}")

        for item in items:
            print(f"  - {item[1]} | {item[2]} | QTY: {item[4]}")

        print(f"Jumlah item: {len(items)}")

        total_items += len(items)

    print(f"\nTOTAL ITEM GROUPING: {total_items}")
    print(f"TOTAL ITEM ASLI: {len(processed)}")

    generated = set()

    for so, items in grouped.items():
        if so in generated:
            continue

        path = generate_po(so, items, SAVE_DIR)
        print(f"✅ PO dibuat: {path}")

        generated.add(so)

# ===== MODE SWITCH =====
    if MODE == "DEBUG":
        print("\n🔄 Export Excel (flat)...")
        output = export_excel(processed, columns, SAVE_DIR)
        print(f"✅ SELESAI (DEBUG): {output}")

    elif MODE == "PROD":
        print("\n🚀 GENERATE PO PER SO...")

    for so, items in grouped.items():
        path = generate_po(so, items, SAVE_DIR)
        print(f"✅ PO dibuat: {path}")

if __name__ == "__main__":
    main()