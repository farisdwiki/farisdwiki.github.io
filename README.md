# farisdwiki.github.io

## Portal Admin Harga Sparepart

Portal admin tersedia melalui menu **Admin** pada website dan dibuka di halaman terpisah `admin.html`.

- Login demo: `admin`
- Password demo: `gunmotor2026`
- Data sementara disimpan di browser menggunakan `localStorage`.
- Struktur data Google Sheet: `id`, `nama`, `stok`, `satuan`, `harga`, `updatedAt`.
- ID baru dibuat dengan pola `SP00001`, `SP00002`, dan seterusnya. Waktu update memakai format `DD/MM/YY HH:MM WIB`.
- Import hanya menerima `.xlsx` dengan struktur yang sama seperti tab Google Sheet `Sparepart`.
- Data yang sudah diinput dapat diperbarui melalui tombol edit pada tabel.

### Menghubungkan Google Sheet

1. Buat Google Sheet baru dengan tab bernama `Sparepart`.
2. Buka **Extensions > Apps Script**, salin isi [google-apps-script.gs](google-apps-script.gs), lalu deploy sebagai **Web app** dengan akses **Anyone**.
3. Salin URL Web app ke variabel `sheetEndpoint` di script portal pada [admin.html](admin.html).
4. Saat dashboard dibuka, data akan dimuat dari Google Sheet. Penambahan, import, dan penghapusan data disinkronkan otomatis setelah jeda singkat.
5. Dashboard mengecek perubahan dari Google Sheet setiap 30 detik. Tombol **Sync Sheet** tetap tersedia untuk sinkronisasi manual.
6. Jika Google Sheet tidak dapat diakses, portal memakai data `localStorage` dan mencoba tersambung kembali secara otomatis.

Catatan: login pada situs static ini adalah login demo sisi klien. Untuk produksi, pindahkan autentikasi ke backend atau gunakan Firebase/Auth0 agar password tidak tersimpan di source code publik.
