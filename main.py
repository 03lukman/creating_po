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

TEMPLATE_PATH = os.getenv("TEMPLATE_PATH")
BASE_PO = os.getenv("BASE_PO_NUMBER", "POR-NN26C023")


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

    # ===== DEBUG VALIDASI =====
    if MODE == "DEBUG":
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

    # ===== MODE SWITCH =====
    if MODE == "DEBUG":
        print("\n🔄 Export Excel (flat)...")
        output = export_excel(processed, columns, SAVE_DIR)
        print(f"✅ SELESAI (DEBUG): {output}")

    elif MODE == "PROD":
        print("\n🚀 GENERATE PO PER SO...")

        success = 0
        skipped = 0

        for i, (so, items) in enumerate(grouped.items()):

            # VALIDASI HARGA
            all_missing = all(item[-2] is None for item in items)

            if all_missing:
                print(f"⛔ SKIP PO {so} (semua harga kosong)")
                skipped += 1
                continue

            path = generate_po(
                so=so,
                items=items,
                save_dir=SAVE_DIR,
                mode=MODE,
                template_path=TEMPLATE_PATH,
                base_po=BASE_PO,
                index=i
            )

            print(f"✅ PO dibuat: {path}")
            success += 1

        print("\n===== SUMMARY =====")
        print(f"Total SO     : {len(grouped)}")
        print(f"PO dibuat    : {success}")
        print(f"PO di-skip   : {skipped}")


if __name__ == "__main__":
    main()