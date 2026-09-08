# farisdwiki.github.io

## Portal Admin Harga Sparepart

Portal admin tersedia melalui menu **Admin** pada website dan dibuka di halaman terpisah `admin.html`.

- Login demo: `admin`
- Password demo: `gunmotor2026`
- Data sementara disimpan di browser menggunakan `localStorage`.
- Import menerima `.xlsx`, `.xls`, dan `.csv` dengan kolom wajib `nama`, `jenis`, `harga`; kolom `merk` dan `stok` opsional.

### Menghubungkan Google Sheet

1. Buat Google Sheet baru dengan tab bernama `Sparepart`.
2. Buka **Extensions > Apps Script**, salin isi [google-apps-script.gs](google-apps-script.gs), lalu deploy sebagai **Web app** dengan akses **Anyone**.
3. Salin URL Web app ke variabel `sheetEndpoint` di script portal pada [admin.html](admin.html).
4. Gunakan tombol **Sync Sheet** setelah data diinput.

Catatan: login pada situs static ini adalah login demo sisi klien. Untuk produksi, pindahkan autentikasi ke backend atau gunakan Firebase/Auth0 agar password tidak tersimpan di source code publik.
