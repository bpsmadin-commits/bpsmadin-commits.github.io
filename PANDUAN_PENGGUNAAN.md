# Panduan Penggunaan Statistik Nilai Ujian

## 1. Siapkan file Excel nilai

Gunakan file `.xlsx` atau `.xls`. Setiap sheet nilai sebaiknya mewakili satu kelas atau satu kelompok kelas.

Pada sheet nilai, buat satu baris header yang berisi minimal:

| NO | NOPES | NAMA | ASR | KLS/KELAS | Mapel 1 | Mapel 2 | Mapel 3 | ... |
|---:|---|---|---|---|---:|---:|---:|---|

Ketentuan penting:

- Header `NAMA` atau `NAMA SISWA` wajib ada.
- Header `NOPES`, `ASR`, dan `KLS`/`KELAS` dibaca otomatis jika tersedia.
- Nama mapel harus berada di baris header yang sama dengan `NAMA`.
- Nilai siswa harus berupa angka. Nilai kosong boleh dibiarkan kosong.
- Jangan memasukkan kolom `JUMLAH`, `RATA-RATA`, `STATUS`, `RANK`, atau rumus hasil lama sebagai mapel.
- Baris judul boleh diletakkan di atas header; web akan mencari baris `NAMA` secara otomatis.
- Jangan menggabungkan sel pada baris header dan data.

Contoh sederhana:

| NO | NOPES | NAMA | ASR | KLS | Fiqih | Akidah | Bahasa Arab |
|---:|---|---|---|---|---:|---:|---:|
| 1 | 00001 | Ahmad Fauzi | A01 | IV J MDTA | 80 | 75 | 90 |
| 2 | 00002 | Siti Aisyah | B02 | IV J MDTA | 70 | 85 | 80 |

## 2. Tambahkan wali kelas

Buat sheet dengan nama persis:

`DATA WALI KELAS`

Gunakan header berikut:

| NO | NAMA | ASR | KELAS | KELAS | JABATAN |
|---:|---|---|---|---|---|
| 1 | Ust. Ahmad Fauzi | B | IV J MDTA | 4 J MDTA | WALI KELAS |

Catatan:

- Kolom `NAMA` berisi nama wali kelas.
- Satu atau beberapa kolom `KELAS` boleh digunakan.
- Isi `JABATAN` dengan `WALI KELAS` agar baris tersebut dibaca.
- Penulisan kelas harus sama atau masih dapat dicocokkan dengan sheet nilai, misalnya `IV J MDTA`, `4 J MDTA`, `IV J MID`, atau `4 J MID`.
- Sheet `DATA WALI KELAS` tidak akan diproses sebagai sheet nilai.
- Jika kelas nilai tidak memiliki pasangan di sheet ini, nama wali kelas tidak akan tampil.

## 3. Upload ke web

1. Buka web Statistik Nilai Ujian.
2. Pilih kuartal yang sesuai.
3. Klik `Upload Data`.
4. Pilih file Excel nilai yang sudah disiapkan.
5. Tunggu sampai daftar kelas dan nilai muncul.
6. Pilih kelas pada filter untuk melihat analisisnya.

File dapat berisi beberapa sheet kelas sekaligus. Nama sheet akan digunakan sebagai identitas kelas jika kolom `KLS`/`KELAS` tidak tersedia.

## 4. Pengaturan KKM

Nilai KKM dapat diubah pada bagian `KKM`. Status siswa akan dihitung otomatis:

- Nilai rata-rata sama dengan atau lebih besar dari KKM: `Tuntas`.
- Nilai rata-rata di bawah KKM: `Tidak Tuntas`.

## 5. Export leger Excel

1. Buka tab `Leger`.
2. Klik `Export Excel Semua Kelas`.
3. Web akan membuat satu file Excel.
4. Setiap kelas dibuat menjadi satu sheet.
5. Rata-rata dihitung dari `JUMLAH ÷ jumlah mapel`.
6. Kolom `JUMLAH`, `RATA-RATA`, `RANKING`, dan `STATUS` berisi rumus Excel.

Jika ingin mengedit nilai setelah export, ubah nilai pada kolom mapel. Rumus jumlah, rata-rata, ranking, dan status akan mengikuti ketika Excel melakukan perhitungan ulang.

## 6. Masalah umum

### Data tidak terbaca

Pastikan ada header `NAMA` atau `NAMA SISWA`, nilai mapel berupa angka, dan baris data berada di bawah header.

### Mapel tidak muncul

Pastikan nama mapel tidak memakai nama kolom yang dicadangkan seperti `NAMA`, `ASR`, `KELAS`, `JUMLAH`, `RATA-RATA`, `STATUS`, atau `RANK`.

### Wali kelas tidak muncul

Pastikan nama sheet adalah `DATA WALI KELAS`, jabatan berisi `WALI KELAS`, dan kelas di sheet tersebut cocok dengan kelas pada sheet nilai.

### Nilai terbaca sebagai kosong

Hapus tanda petik, teks, atau simbol dari nilai. Gunakan angka biasa seperti `80`, `75.5`, atau `100`.
