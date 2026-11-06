# Pengambilan Data Bahan Baku

Modul browser-side yang hanya membaca sheet `HEADER` dan `BARANG` dari file CEISA `.xlsx`. Sheet lain tidak digunakan sebagai sumber.

## Aturan utama
- Mendukung kode dokumen `27`, `23`, dan `40`.
- BC 2.7 / BC 2.3: NDPBM dari HEADER kolom BW; kode asal bahan baku `0`; BAHANBAKUTARIF membentuk BM, PPN, PPH.
- BC 4.0: NDPBM `1`; kode asal bahan baku `1`; Harga Penyerahan dari BARANG kolom AH; BAHANBAKUTARIF membentuk PPNLOKAL.
- BAHANBAKU memakai 41 kolom dan BAHANBAKUTARIF memakai 25 kolom mengikuti workbook contoh.
- NETTO, BRUTO, NILAI BAYAR, NILAI FASILITAS, NILAI SUDAH DILUNASI dan kolom yang tidak dipetakan dibiarkan kosong. VOLUME diisi CIF per satuan untuk BC 2.7/2.3, serta Harga Penyerahan per satuan untuk BC 4.0; pembaginya memakai JUMLAH SATUAN asli dari CEISA sehingga VOLUME tidak berubah ketika user mengedit JUMLAH SATUAN.
- Jumlah satuan, seri barang, seri bahan baku, dan tarif dapat diedit sebelum disalin/diunduh.
- Pemilihan barang disinkronkan untuk kedua output.

## Catatan validasi
Kode PPNLOKAL dan tarif default mengikuti aturan kerja yang disepakati pada pengembangan awal; tetap perlu diuji dengan contoh transaksi BC 4.0 yang sudah dipakai operasional.
