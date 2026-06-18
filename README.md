# Scraping & Data Cleaning Info Kost Jogja

Project ini dibuat untuk mengolah data postingan info kost di Jogja dari teks mentah menjadi data yang lebih rapi dalam bentuk tabel CSV dan Excel.

Project ini tidak menggunakan bot otomatis untuk mengambil data dari Facebook. Data postingan dikumpulkan secara manual dari postingan publik, lalu diproses menggunakan Python.

## Tujuan Project

Tujuan utama project ini adalah membantu mencari informasi kost dengan kriteria tertentu, yaitu:

* Harga 650 sampai 750 ribu
* Kamar mandi dalam
* Kost putra atau campur
* Area Maguwoharjo, Condongcatur, atau Jalan Kaliurang
* Hanya postingan penawaran kost, bukan postingan pencarian

## Fitur

* Membaca data postingan mentah dari file `.txt`
* Memisahkan setiap postingan berdasarkan tanda `---`
* Mengambil informasi harga kost
* Mengambil area atau lokasi kost
* Mengambil tipe kost
* Mengambil status kamar mandi
* Mengambil fasilitas kost
* Membedakan postingan penawaran dan postingan pencarian
* Melakukan filter data sesuai kriteria
* Export hasil data ke CSV dan Excel
* Membuat ringkasan analisis sederhana

## Struktur Folder

```text
scraping-kost-jogja/
│
├── data/
│   ├── raw/
│   │   └── postingan_kost.txt
│   └── clean/
│       ├── data_kost_jogja.csv
│       ├── data_kost_jogja.xlsx
│       ├── data_kost_jogja_filter.csv
│       └── data_kost_jogja_filter.xlsx
│
├── output/
│   └── ringkasan_kost.txt
│
├── src/
│   ├── cleaner.py
│   ├── analisis.py
│   └── export_excel.py
│
├── requirements.txt
├── .gitignore
└── README.md
```

## Teknologi yang Digunakan

* Python
* pandas
* openpyxl
* regular expression / regex

## Cara Menjalankan Project

### 1. Clone atau buka folder project

```bash
cd scraping-kost-jogja
```

### 2. Aktifkan virtual environment

Untuk Windows:

```bash
venv/Scripts/activate
```

### 3. Install library

```bash
pip install -r requirements.txt
```

### 4. Masukkan data mentah

Masukkan data postingan ke file berikut:

```text
data/raw/postingan_kost.txt
```

Pisahkan setiap postingan dengan tanda:

```text
---
```

Contoh format data:

```text
Kos putra daerah Maguwoharjo. Harga 700rb/bulan. Kamar mandi dalam, WiFi, parkir motor.

---

Kos campur daerah Condongcatur. Harga 750rb/bulan. Kamar mandi dalam.
```

### 5. Jalankan proses cleaning dan filter

```bash
python src\cleaner.py
```

Hasil file akan tersimpan di:

```text
data/clean/data_kost_jogja.xlsx
data/clean/data_kost_jogja_filter.xlsx
```

### 6. Jalankan analisis ringkas

```bash
python src\analisis.py
```

Hasil ringkasan akan tersimpan di:

```text
output/ringkasan_kost.txt
```

## Output

Project ini menghasilkan dua jenis data.

### 1. Semua Data

Berisi semua postingan yang berhasil dibaca dari file mentah.

File:

```text
data/clean/data_kost_jogja.xlsx
```

### 2. Data Filter

Berisi data kost yang sesuai dengan kriteria pencarian.

File:

```text
data/clean/data_kost_jogja_filter.xlsx
```

## Contoh Hasil Filter

| Area        | Tipe Kost   | Harga | Kamar Mandi |
| ----------- | ----------- | ----: | ----------- |
| Maguwoharjo | Putra       |   650 | Dalam       |
| Maguwoharjo | Putra       |   750 | Dalam       |
| Maguwoharjo | Putra/Putri |   700 | Dalam       |

## Catatan Etika Data

Project ini tidak mengambil data pribadi seperti nama akun, nomor HP, foto orang, komentar, atau isi chat pribadi. Data yang digunakan hanya teks informasi kost yang relevan untuk analisis harga, lokasi, tipe kost, dan fasilitas.

## Status Project

Versi awal project sudah berjalan dan dapat digunakan untuk membersihkan serta memfilter data postingan info kost Jogja.
