from pathlib import Path
import pandas as pd

FILE_FILTER = Path("data/clean/data_kost_jogja_filter.xlsx")
OUTPUT_RINGKASAN = Path("output/ringkasan_kost.txt")


def baca_data_filter():
    if not FILE_FILTER.exists():
        print("File filter belum ditemukan.")
        print("Jalankan dulu: python src\\cleaner.py")
        return None

    df = pd.read_excel(FILE_FILTER)
    return df


def buat_ringkasan(df):
    if df.empty:
        return "Belum ada data kost yang sesuai kriteria."

    total_data = len(df)
    harga_termurah = df["harga"].min()
    harga_termahal = df["harga"].max()
    rata_rata_harga = df["harga"].mean()

    area_terbanyak = df["area"].value_counts().idxmax()
    tipe_terbanyak = df["tipe_kost"].value_counts().idxmax()

    ringkasan = f"""
RINGKASAN DATA KOST JOGJA

Total kost sesuai kriteria : {total_data}
Harga termurah             : {harga_termurah} ribu
Harga termahal             : {harga_termahal} ribu
Rata-rata harga            : {rata_rata_harga:.0f} ribu
Area paling banyak         : {area_terbanyak}
Tipe kost paling banyak    : {tipe_terbanyak}

Kriteria filter:
- Harga 650 sampai 750 ribu
- Kamar mandi dalam
- Kost putra atau campur
- Area Maguwoharjo, Condongcatur, atau Jalan Kaliurang
- Hanya postingan penawaran, bukan postingan pencarian
"""

    return ringkasan.strip()


def simpan_ringkasan(ringkasan):
    OUTPUT_RINGKASAN.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_RINGKASAN.write_text(ringkasan, encoding="utf-8")

    print(ringkasan)
    print("\nRingkasan berhasil disimpan ke:")
    print(OUTPUT_RINGKASAN)


if __name__ == "__main__":
    df = baca_data_filter()

    if df is not None:
        ringkasan = buat_ringkasan(df)
        simpan_ringkasan(ringkasan)