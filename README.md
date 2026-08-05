# SiRUP — Sanding RKA & RUP

Userscript untuk menarik data **RKA** dan **RUP** langsung dari
[sirup.inaproc.id](https://sirup.inaproc.id) lalu mengekspornya menjadi satu file
Excel penyandingan: berapa pagu yang wajib diumumkan, berapa yang sudah masuk RUP,
dan bagian mana yang masih menganga.

Dibuat untuk keperluan monitoring & evaluasi internal satker.

---

## Instalasi

1. Pasang [Tampermonkey](https://www.tampermonkey.net/) di Chrome/Edge/Firefox.
2. Buka [`sirup_exporter.user.js`](https://raw.githubusercontent.com/Fakhry-Glob/sirup-sanding-rka-rup/master/sirup_exporter.user.js)
   — Tampermonkey otomatis menawarkan pemasangan.
3. Login ke SiRUP, buka halaman mana pun di `sirup.inaproc.id/sirup/*`.
4. Tombol **Ekspor Sanding RKA & RUP** muncul di kanan bawah.

Pembaruan berikutnya ditarik otomatis oleh Tampermonkey dari branch `master`.

## Cara pakai

1. Klik tombol di kanan bawah. Panel akan menampilkan **tahun anggaran** dan
   **nama satker** yang terdeteksi dari halaman — koreksi kalau meleset.
2. Klik **Mulai Ekspor**. Panel progres menampilkan tahapan
   (Persiapan → Tarik RKA → Tarik RUP → Menyanding → Susun Excel), persentase,
   waktu berjalan, dan log. Proses bisa dihentikan lewat tombol **Hentikan**.
3. File `Sanding_RKA_RUP_<Satker>_TA<tahun>_<tanggal>.xlsx` otomatis terunduh.

Untuk satker besar prosesnya bisa beberapa menit — biarkan tab tetap terbuka.
Semua permintaan memakai sesi login yang sedang berjalan; tidak ada kredensial
yang disimpan atau dikirim ke mana pun.

## Isi file Excel

| Sheet | Isi |
|---|---|
| **Dashboard Evaluasi** | Kartu KPI (pagu satker, non-pengadaan, target pengadaan, RUP terumumkan), capaian, statistik paket, 10 komponen dengan selisih terbesar, legenda warna |
| **Ringkasan Pagu** | Daftar seluruh komponen RKA beserta pagunya |
| **Sanding RKA & RUP** | Inti laporan: per komponen — pagu RKA, non-pengadaan, target pengadaan, RUP terumumkan, selisih, capaian, status. Lengkap dengan subtotal RO → KRO → Kegiatan → Program |
| **Daftar Paket RUP** | Semua paket penyedia beserta status A / FD / U dan komponen RKA yang tersanding |
| **Paket RUP Tanpa Sandingan** | Paket yang MAK-nya tidak ketemu di RKA — biasanya salah input atau anggaran sudah dihapus |
| **Detail – \<kode program\>** | Rincian RKA sampai level detail, disandingkan baris per baris dengan paket RUP-nya |

Semua sheet sudah memakai freeze pane, autofilter, format Rupiah, dan setelan
cetak A4 fit-to-width dengan header berulang.

## Cara penyandingan bekerja

- **Kunci penyandingan** adalah MAK 7 ruas: `program.kegiatan.KRO.RO.komponen.subkomponen.akun`.
  MAK dari RUP dinormalkan dulu (bentuk 9 ruas berawalan tahun & kode satker juga diterima).
- **Non-pengadaan** (centang NP / Gaji) dikeluarkan dari target. Centang di level
  komponen, sub-komponen, atau akun **diturunkan** ke seluruh baris detail di bawahnya.
- **Target pengadaan** = pagu RKA − belanja non-pengadaan.
- Hanya paket berstatus **A + FD + U** (draft PPK, final draft, dan sudah
  diumumkan KPA) yang dihitung sebagai realisasi RUP. Paket draft/batal tetap
  ditampilkan sebagai peringatan berwarna oranye.
- Satu paket RUP yang menutup beberapa baris RKA dialokasikan **proporsional**;
  sisa alokasi jatuh ke baris pertama pada kunci akun yang sama.

## Batasan yang diketahui

- **Paket swakelola belum ikut disandingkan.** Jumlahnya dilaporkan sebagai
  catatan di log, tapi tidak masuk perhitungan.
- Butuh sesi login SiRUP yang aktif; script tidak melakukan autentikasi sendiri.
- Deteksi tahun anggaran & nama satker mengandalkan struktur halaman SiRUP. Kalau
  SiRUP mengubah markup-nya, nilai itu bisa meleset — makanya keduanya bisa
  dikoreksi manual di panel sebelum ekspor.
- Pengambilan detail paket dibatasi 1.000 paket per satker per tahun.

## Pengembangan

File yang dipakai hanya satu: `sirup_exporter.user.js`. Tidak ada proses build.

Cek sintaks:

```bash
node --check sirup_exporter.user.js
```

Untuk menguji pembentukan Excel tanpa membuka SiRUP, jalankan `buildExcel()` di
Node dengan ExcelJS asli dan data sintetis (stub `document`, `Blob`, dan `saveAs`),
lalu baca ulang file hasilnya untuk memeriksa `numFmt`, warna, freeze pane, dan
autofilter.

> Catatan: properti format angka di ExcelJS adalah `cell.numFmt` — bukan
> `numFormat`. Salah nama tidak memunculkan error apa pun, formatnya hanya diam-diam
> tidak terpasang. Warna wajib ARGB 8 digit (`FF` + RRGGBB); nilai 6 digit dibaca
> salah oleh Excel.

Skrip Python satu kali pakai dari fase awal proyek disimpan di `legacy/` sebagai
arsip — jalurnya masih menunjuk ke mesin penulis aslinya dan tidak dipakai lagi.
