# Gun Motor Madiun - Sistem Informasi & Manajemen Sparepart

WebApps ini merupakan sistem informasi berbasis web untuk **Gun Motor**, sebuah bengkel mobil profesional di Madiun. Sistem ini dirancang untuk memberikan informasi layanan kepada pelanggan sekaligus memudahkan pengelola bengkel dalam memanajemen inventaris dan harga sparepart.

Sistem ini terdiri dari dua bagian utama:
1. **Website Publik (Landing Page)**: Halaman interaktif yang menyajikan profil bengkel, jenis layanan, paket perawatan berkala, galeri fasilitas, dan kemudahan akses konsultasi langsung via WhatsApp bagi pelanggan.
2. **Portal Admin (Dashboard Sparepart)**: Halaman *single-page application* (SPA) khusus pengelola untuk mendata, menambah, mengedit, menghapus, serta melakukan import data (via file `.xlsx`) harga dan stok sparepart. Portal ini dilengkapi dengan fitur pencarian dan paginasi data yang responsif.

## Arsitektur & Database (Google Sheets)

Aplikasi ini bersifat *serverless* dan menggunakan **Google Sheets** sebagai basis datanya. Backend dioperasikan melalui **Google Apps Script** (GAS) yang berfungsi sebagai REST API untuk membaca dan menulis data dari Portal Admin.

- Jika koneksi internet terputus, portal menggunakan penyimpanan lokal (`localStorage`) pada browser sebagai *cache* sementara, dan akan melakukan sinkronisasi otomatis (*auto-sync*) dengan Google Sheets ketika koneksi kembali tersedia.
- Struktur kolom data pada Google Sheets meliputi: `id`, `nama`, `stok`, `satuan`, `harga`, dan `updatedAt`.
- Waktu pembaharuan (`updatedAt`) selalu direkam secara otomatis dengan format string lokal yang jelas untuk menjaga integritas catatan waktu setiap perubahan data.

## Panduan Setup Database

Untuk menghubungkan Portal Admin dengan Google Sheets Anda sendiri, ikuti langkah-langkah berikut:

1. Buat **Google Sheet** baru dan pastikan tab pertamanya bernama `Sparepart`.
2. Pada menu Google Sheets, buka **Extensions > Apps Script**.
3. Salin seluruh kode yang ada di dalam file `google-apps-script.gs` ke dalam editor Apps Script tersebut.
4. Lakukan deployment melalui menu **Deploy > New deployment**. Pilih tipe **Web app**. Atur akses ke **"Anyone"**.
5. Salin **URL Web app** yang diberikan, lalu *paste* (ganti nilai) ke dalam variabel `sheetEndpoint` yang ada di dalam file `admin.html`.
6. Buka Portal Admin. Sistem akan otomatis terhubung, dan Anda dapat mulai melakukan sinkronisasi data dari dan ke Google Sheets.

> **Catatan Keamanan:** Untuk lingkungan produksi publik, disarankan menggunakan sistem autentikasi yang mumpuni (seperti Firebase Auth atau Auth0) untuk melindungi rute menuju dashboard admin.
