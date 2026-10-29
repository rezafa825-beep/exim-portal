# EXIM WhatsApp Report Generator (Versi Awal)

## Isi
- `index.html` — struktur halaman
- `style.css` — tampilan
- `app.js` — pembacaan Excel, filter, jalur, dan pembuat laporan

## Cara menjalankan di komputer
1. Ekstrak ZIP.
2. Buka `index.html` dengan Chrome/Edge.
3. Pastikan koneksi internet tersedia agar library SheetJS dapat dimuat dari CDN.
4. Pilih file Excel CEISA. File hanya dibaca di browser; jangan unggah file perusahaan ke GitHub.
5. Pilih jenis dokumen, Customer/Supplier, nama pihak, isi Jenis Barang, dan cek nomor pendaftaran yang perlu dijalur MERAH.
6. Klik **Buat / Perbarui Laporan**, lalu **Salin Laporan**.

## Cara mengunggah ke GitHub Pages
1. Buka repository GitHub yang akan digunakan.
2. Unggah `index.html`, `style.css`, dan `app.js` ke root repository (bukan file Excel).
3. Buka Settings → Pages.
4. Pilih Deploy from a branch, pilih `main` dan folder `/ (root)`, lalu Save.
5. Tunggu URL Pages aktif.

## Catatan versi awal
- Nomor daftar dibaca dari HEADER kolom CP; tanggal daftar dari kolom CQ; tujuan transaksi dari kolom N.
- Kode dokumen 27 dibedakan masuk/keluar memakai ENTITAS kode 3 seperti kesepakatan.
- Pihak keluar memakai ENTITAS kode 8; pihak masuk memakai kode 9. Mode Customer/Supplier dapat dipilih manual.
- Semua dokumen default HIJAU; centang nomor daftar untuk mengubahnya menjadi MERAH.
- Jenis barang masih diisi manual karena nama barang pada BARANG bisa berupa uraian panjang dan belum ada aturan kategorisasi yang disepakati.
- Pastikan hasil nomor pendaftaran, kuantitas, kemasan, segel, dan tujuan transaksi diverifikasi sebelum digunakan.
- Library SheetJS dimuat dari CDN, sehingga perlu koneksi internet ketika membuka web.
