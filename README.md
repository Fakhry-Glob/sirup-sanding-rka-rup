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
| **Dashboard Evaluasi** | Kartu KPI (pagu satker, non-pengadaan, target pengadaan, RUP terumumkan), capaian, statistik paket, kontrol silang total RUP, 10 komponen dengan selisih terbesar, legenda warna |
| **Ringkasan Pagu** | Daftar seluruh komponen RKA beserta pagunya |
| **Sanding RKA & RUP** | Inti laporan: per komponen — pagu RKA, non-pengadaan, target pengadaan, RUP terumumkan, selisih, capaian, status. Lengkap dengan subtotal RO → KRO → Kegiatan → Program |
| **Daftar Paket RUP** | Semua paket penyedia & swakelola beserta status A / FD / U, komponen RKA tersanding, dan **Status Sanding** per paket (penuh / sebagian berapa persen / tidak sama sekali) lengkap dengan nilai yang tertampung dan yang tidak |
| **Paket RUP Tanpa Sandingan** | Paket yang MAK-nya tidak ketemu di RKA — biasanya salah input atau anggaran sudah dihapus |
| **Detail – \<kode program\>** | Rincian RKA sampai level detail, disandingkan baris per baris dengan paket RUP-nya |

Semua sheet sudah memakai freeze pane, autofilter, format Rupiah, dan setelan
cetak A4 fit-to-width dengan header berulang.

## Cara penyandingan bekerja

- **Kunci penyandingan** adalah MAK 7 ruas: `program.kegiatan.KRO.RO.komponen.subkomponen.akun`.
  MAK dari RUP dinormalkan dulu. SiRUP memakai dua bentuk, dan **panjangnya tidak
  bisa dipakai membedakan** — awalan `tahun.kode satker` dikenali dari ruas pertama
  yang berupa tahun 4 digit, bukan dari jumlah ruas:

  | Sumber | Contoh nyata | Hasil |
  |---|---|---|
  | Tabel detail paket | `WA.2378.EBA.994.002.AD.521811` | 7 ruas polos, dipakai menyanding |
  | Kolom daftar paket | `2026.14564.WA.2378.EBA.994.002.AD.521811` | awalan dibuang → 7 ruas |
  | Kolom daftar paket | `2026.14564.DL.2376.FAN.ZZ1.ZZ1` | awalan dibuang → sisa 5 ruas: **MAK belum lengkap** |

  Bentuk ketiga itu juga 7 ruas. Sebelum v2.3 awalannya tidak dibuang, jadi program
  terbaca `2026` dan kegiatan `14564` — paketnya dilaporkan "anggaran dihapus atau
  salah input", padahal sebenarnya MAK-nya baru terisi sampai level KRO/RO.
  Sekarang sebabnya disebut apa adanya di sheet *Paket RUP Tanpa Sandingan*.

  `ZZ1` adalah penanda SAKTI untuk RO/komponen yang belum ditentukan.
- **Non-pengadaan** (centang NP / Gaji) dikeluarkan dari target. Centang di level
  komponen, sub-komponen, atau akun **diturunkan** ke seluruh baris detail di bawahnya.
- **Target pengadaan** = pagu RKA − belanja non-pengadaan.
- Hanya paket berstatus **A + FD + U** (draft PPK, final draft, dan sudah
  diumumkan KPA) yang dihitung sebagai realisasi RUP. Paket draft/batal tetap
  ditampilkan sebagai peringatan berwarna oranye.
- **Paket penyedia dan swakelola dua-duanya ikut disanding** sejak v3.0, dengan
  aturan A+FD+U yang sama. Kalau tidak ada satu pun swakelola yang lolos A+FD+U,
  log menyebutnya supaya kolom A di SiRUP bisa diperiksa dulu.
### Alokasi paket ke baris RKA

Baris RKA diisi berurutan; tiap baris menyerap `min(sisa pagu baris, sisa pagu
paket)`, dan paket yang masih bersisa mengalir ke baris berikutnya. Satu paket
boleh menutup beberapa baris, dan **tiap baris menampilkan porsinya sendiri**:

| Baris RKA | Pagu RKA | Paket | Porsi |
|---|---:|---|---:|
| Pagu A | 30.000.000 | Paket A1 | 12.000.000 |
|  |  | Paket A2 | 18.000.000 |
| Pagu B | 20.000.000 | Paket B1 | 7.000.000 |
|  |  | Paket B2 | 13.000.000 |
| Pagu C | 10.000.000 | Paket B2 | 10.000.000 |

Paket B2 berpagu 23.000.000 dan terbelah ke dua baris. Urutannya tiga tahap:

1. **Cocok persis 7 ruas** — MAK paket sama dengan MAK baris RKA.
2. **Luber lintas akun, dalam komponen yang sama.** Satu kontrak outsourcing
   lazimnya menutup honor, jaminan ketenagakerjaan, dan jaminan kesehatan
   sekaligus — tiga akun berbeda. Baris yang terisi dari tahap ini ditandai
   *"[sebagian dari paket berakun lain]"* pada kolom nama paket.
3. **Sisa yang tidak tertampung** mendapat barisnya sendiri, lengkap dengan ID
   paketnya, dipisah menurut sebabnya.

### Tiga jenis temuan pada baris sisa

| Warna | Kapan muncul | Yang perlu diperiksa |
|---|---|---|
| 🔴 Merah — **TEMUAN** | Ada paket RUP di atas pagu yang ditandai **NP/Gaji** | Tanda NP/Gaji-nya keliru, pagunya sudah direvisi sehingga baris itu berubah sifat, atau RUP-nya belum disesuaikan |
| 🟠 Oranye | RUP **melebihi pagu** RKA | Pagu sudah direvisi turun, RUP belum disesuaikan, atau nilai paketnya memang melebihi pagu |
| 🟠 Oranye | **MAK tidak ada** di RKA komponen ini | Penulisan MAK pada paket, atau akunnya hilang/berubah saat revisi |

Paket yang MAK-nya menunjuk akun ber-tanda NP **sengaja ditahan** dari tahap 2.
Kalau dibiarkan meluber ke akun lain, paket itu akan terlihat tersanding rapi
dan temuannya hilang — padahal justru itu yang perlu dikoreksi satker.

**Bukti revisi ikut disebut.** Kolom *pagu sebelum revisi* sudah ditarik dari
RKA, jadi kalau pagu di akun itu memang berubah, barisnya menambahkan:

> Catatan: pagu baris di akun ini BERUBAH saat revisi (sebelum revisi
> Rp 25.000.000) — kemungkinan besar RUP belum disesuaikan.

atau, untuk baris yang belum ada sebelum revisi, *"baris di akun ini BARU muncul
setelah revisi"*. Kalau tidak ada jejak revisi, catatan ini tidak dimunculkan —
daripada menebak.

### Seluruhnya atau sebagian?

Ketidaksinkronan jarang bersifat "paket ini salah". Ada tiga tingkat, dan
laporan menyebut ketiganya:

- **Paket** bisa punya beberapa baris MAK — sebagian cocok, sebagian tidak.
- **Satu baris MAK** pun bisa terbelah: sebagian terserap baris RKA, sisanya
  tidak tertampung.
- **Akun** bisa kelebihan muatan karena beberapa paket menunjuk MAK yang sama.

Karena itu sheet *Daftar Paket RUP* punya kolom **Status Sanding** per paket:
*Tersanding penuh* / *Tersanding SEBAGIAN — 77% (n baris MAK)* / *Tidak
tersanding sama sekali*, ditambah kolom nilai yang tertampung dan yang tidak.
Baris kelebihan di sheet Detail juga merinci nominal **per paket**, bukan satu
angka gabungan untuk beberapa ID sekaligus.

**Sebuah baris tidak pernah menerima lebih dari pagunya.** Sampai v2.3 seluruh
sisa paket ditumpahkan ke baris pertama tiap akun, sehingga baris berpagu
45.500.000 bisa tampil 45.500.000 → 377.999.808 dan kolom Selisih-nya tidak
berarti apa-apa. Subtotal sub-komponen & akun ikut memakai angka hasil alokasi,
bukan jumlah menurut MAK, supaya baris detail selalu berjumlah sama dengan
subtotal di atasnya.

Angka di **Dashboard** dan **Sanding RKA & RUP** tidak terpengaruh perubahan ini
— keduanya dijumlah langsung dari baris MAK, bukan dari hasil alokasi.

## Kontrol silang angka RUP

Dashboard menghitung total pagu RUP terumumkan dengan tiga cara: (A) dari daftar
paket, (B) dari baris MAK hasil penarikan detail paket, (C) setelah disandingkan
ke komponen RKA. Ketiganya seharusnya sama.

- **B atau C lebih besar dari A** → ada pagu terhitung ganda; blok itu ditandai
  merah beserta dugaan tahap penyebabnya, dan capaian di kartu KPI tidak bisa
  dipakai apa adanya.
- **A lebih besar dari B** → wajar, itu paket yang MAK-nya tidak terbaca;
  daftarnya ada di sheet *Paket RUP Tanpa Sandingan*.

## Penjaga sebelum & selama penarikan

Beberapa hal dulu bisa lolos diam-diam dan menghasilkan laporan yang tampak wajar
tapi salah. Sekarang masing-masing punya penjaga sendiri:

| Penjaga | Kapan berbunyi |
|---|---|
| **Tahun tidak cocok** — ekspor dihentikan | Halaman RKA selalu mengikuti tahun aktif **sesi** SiRUP; dropdown di panel hanya menggerakkan sisi RUP. Kalau keduanya beda, laporan akan menyanding dua tahun berlainan dan tetap diberi label tahun panel. Ganti tahun aktif lewat menu SiRUP, bukan lewat panel. |
| **Kode satker pada MAK** | Awalan MAK memuat kode satker pemilik pagu. Kalau paket berasal dari lebih dari satu satker, itu disebut — parameter `idSatker` yang dikirim ke SiRUP hanya tebakan dan bisa diabaikan server, sementara nama satker di kop diketik manual. |
| **Kunci komponen bertabrakan** | Dua komponen menghasilkan `program.kegiatan.KRO.RO.komponen` yang sama, atau ada kode yang kosong. Pagu RUP pada kunci itu terhitung ganda. |
| **Paket kembar** | Nama, MAK, dan pagu sama persis — biasanya entri ganda. |
| **Tabel pendanaan tidak dikenali** | Menyebut berapa dari sekian paket yang gagal dibaca. Kalau angkanya besar, struktur halaman SiRUP berubah dan laporannya tidak bisa dipakai. |
| **Penarikan gagal sebagian** | Jumlah tabel komponen RKA atau detail paket yang gagal ditarik. |

## Batasan yang diketahui

- Butuh sesi login SiRUP yang aktif; script tidak melakukan autentikasi sendiri.
  Kalau sesi habis di tengah jalan, proses berhenti dengan pesan jelas — bukan
  menghasilkan laporan berisi nol.
- **Pagu blokir / catatan halaman IV DIPA tidak terlihat** dari data yang ditarik.
  Pagu yang diblokir tetap dihitung sebagai target pengadaan, jadi selisihnya
  bisa wajar tanpa bisa dibedakan otomatis.
- Laporan ini **potret saat diekspor**. Status paket bisa berubah setelahnya.
- Deteksi tahun anggaran & nama satker mengandalkan struktur halaman SiRUP. Kalau
  SiRUP mengubah markup-nya, nilai itu bisa meleset — makanya keduanya bisa
  dikoreksi manual di panel sebelum ekspor.
- Daftar paket ditarik per halaman 1.000 sampai habis, dengan pagar 20.000 paket
  per satker per tahun. Kalau pagar itu kena, log menandainya sebagai peringatan
  merah.

## Kalau ada yang gagal ditarik

Kegagalan jaringan per paket tidak lagi lewat diam-diam. Panel log memunculkan
peringatan merah dan menyebut berapa banyak yang gagal:

- **tabel komponen RKA gagal ditarik** → rincian detail komponen itu kosong,
  jadi non-pengadaan dan sandingan detailnya ikut salah;
- **detail paket RUP gagal ditarik** → pagu paket itu tidak masuk realisasi,
  sehingga capaian terlihat lebih rendah dari seharusnya.

Dua-duanya berarti hal yang sama: ulangi ekspor sebelum angkanya dipakai untuk
monev.

## Pengembangan

File yang dipakai hanya satu: `sirup_exporter.user.js`. Tidak ada proses build.

Cek sintaks:

```bash
node --check sirup_exporter.user.js
```

### Harness uji tanpa SiRUP

`dev/harness.html` adalah halaman SiRUP palsu: seluruh permintaan jaringan
di-stub, jadi panel, progres, penomoran halaman paket, dan penanganan kegagalan
bisa diuji tanpa login dan tanpa membebani server SiRUP.

```bash
python -m http.server 4319
# buka http://localhost:4319/dev/harness.html
```

Yang sengaja disiapkan di dalamnya:

- **1.200 paket** — di atas satu halaman, jadi penarikan bertahap ikut teruji;
- **4 paket yang dibalas HTTP 503** — memunculkan lencana peringatan dan
  catatan "data tidak utuh" di akhir;
- **sel NP berupa ikon glyphicon**, bukan `<input>` — bentuk markup yang dulu
  membuat non-pengadaan terbaca nol;
- **nama satker berisi `& " < >`** — menangkap kebocoran escape di panel;
- `saveAs` dicegat, jadi tidak ada file yang benar-benar terunduh.

Ubah `TOTAL_PACKETS` dan `FAILING_DETAILS` di bagian atas file untuk menguji
skenario lain.

`dev/harness-alokasi.html` khusus menguji alokasi paket ke baris RKA dengan data
kecil yang bisa dihitung tangan (kasus Pagu A/B/C di atas, termasuk Paket B2 yang
terbelah lintas akun). Halaman ini membaca ulang file Excel hasilnya dan menaruh
isi sheet Detail di `window.__detail`, jadi angkanya bisa diperiksa dari console.

Tiap jenis temuan punya sakelarnya sendiri, bisa digabung:

| Query | Yang diuji |
|---|---|
| *(tanpa query)* | Kasus Pagu A/B/C — semua Selisih harus 0 |
| `?excess=1` | Paket 5.000.000 yang tidak muat di mana pun |
| `?np=1` | Paket dibuat di atas pagu ber-tanda NP |
| `?nokey=1` | MAK menunjuk akun yang tidak ada di RKA |
| `?swa=1` | Paket swakelola — kolom daftar tertukar, tabel pendanaan 5 kolom & ter-nest |
| `?kembar=1` | Dua paket identik (nama, MAK, pagu sama persis) |
| `?bentrok=1` | Dua komponen berkode sama di bawah satu RO |
| `?tahunbeda=1` | RKA sesi TA 2025 sementara panel TA 2026 — ekspor harus berhenti |
| `?sebagian=1` | Satu paket yang hanya sebagian pagunya tertampung |

Baris *Pagu A* sengaja dibuat berubah saat revisi (25.000.000 → 30.000.000)
supaya catatan bukti revisi ikut teruji.

Untuk menguji pembentukan Excel tanpa membuka SiRUP, jalankan `buildExcel()` di
Node dengan ExcelJS asli dan data sintetis (stub `document`, `Blob`, dan `saveAs`),
lalu baca ulang file hasilnya untuk memeriksa `numFmt`, warna, freeze pane, dan
autofilter.

Membaca ulang file dengan ExcelJS **tidak cukup** — ExcelJS jauh lebih longgar
daripada Excel. Periksa juga XML mentah di dalam `.xlsx` (unzip, lalu cek
`xl/worksheets/sheet*.xml`), karena beberapa kesalahan hanya ketahuan sebagai
"Workbook Repaired" saat dibuka Excel.

> Tiga jebakan yang pernah menggigit di repo ini:
>
> - Properti format angka adalah `cell.numFmt`, **bukan** `numFormat`. Salah nama
>   tidak memunculkan error apa pun — formatnya diam-diam tidak terpasang.
> - Warna wajib ARGB 8 digit (`FF` + RRGGBB); nilai 6 digit dibaca salah oleh Excel.
> - `views: [{ state: 'frozen', xSplit: 0, ySplit: 0 }]` menghasilkan elemen
>   `<pane>` tidak sah. Sheet tanpa titik beku harus memakai view biasa.

Skrip Python satu kali pakai dari fase awal proyek disimpan di `legacy/` sebagai
arsip — jalurnya masih menunjuk ke mesin penulis aslinya dan tidak dipakai lagi.
