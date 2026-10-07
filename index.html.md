# Auction Result List Checker

Website untuk deploy & cek otomatis file Excel "Result List" lelang.

## Fitur Pengecekan

1. **Cek Tahun Unit dari Nomor Rangka**
   - Mengambil digit ke-10 nomor rangka (standar VIN)
   - Mapping kode: R=2024, S=2025, T=2026, dst.
   - Bandingkan dengan kolom `TAHUN`

2. **Cek Tipe Unit dari CC vs Merk/Type**
   - Ekstrak angka CC dari kolom `MERK/TYPE` (mis. "VARIO CBS 160" → 160)
   - Bandingkan dengan kolom `CC` (toleransi ±2 CC)

3. **Cek Nama STNK vs BPKB**
   - Bandingkan kolom `Nama BPKB` dengan `NAMA STNK`

4. **Cek Note 1 "AN PT"**
   - Jika nama BPKB/STNK diawali "PT", maka Note 1 wajib mengandung "AN PT"

## Cara Pakai

1. Buka `index.html` di browser (atau deploy ke hosting statis)
2. Upload file `.xlsx`
3. Lihat ringkasan & tabel hasil pengecekan
4. Filter berdasarkan jenis masalah
5. Export hasil ke Excel

## Deploy

Cukup upload folder ke:
- **Netlify / Vercel**: drag & drop folder
- **GitHub Pages**: push ke repo, aktifkan Pages
- **Hosting biasa**: upload via FTP

Tidak butuh backend — semua proses di browser.