# Catatan v2: slide 5, 6, 13 (di luar PPT)

- **Deck:** `D:\Download\Financial Analysis SIG (SMGR) v2.pptx`, dibangun dari v1 yang tidak disentuh.
- **Build ulang:** `python scripts\build_v2.py`. Semua angka dibaca langsung dari `WP - Financial Analysis SIG (SMGR) v0.xlsx`; tidak ada angka yang diketik manual kecuali yang ditandai *web* di bawah.

## 1. Isu definisi yang ditunda (perlu diputuskan)

| # | Isu | Posisi di v2 | Opsi |
|---|---|---|---|
| 1 | "COGS" di chart & headline kiri slide 5 = **COGS + ongkos angkut** (FY21 27.350 vs COGS murni 24.976) | **Diputuskan user (18 Okt 2026): tetap COGS + angkut**, supaya sebasis dengan INTP (angkut dijumlahkan ke COGS di kedua perusahaan) | – |
| 2 | Net profit SIG = **laba ke pemilik entitas induk** (2.047 → 191), sementara slide 20 memakai laba konsolidasi (2.117 → 179) | slide 5 sudah diberi note "profit attributable to owners" | samakan basis di seluruh deck |
| 3 | Net margin INTP FY25: **8,9%** (tanpa one-off divestasi Rp670 M) vs **12,7%** di sheet `Ratios` | v2 memakai 8,9% (sama dengan `Ratios v2`) | hapus angka 12,7% di sheet `Ratios` |
| 4 | D/E dan rasio utang di slide 6 memakai definisi `Ratios v2` (utang berbunga termasuk syirkah & sewa; ekuitas = porsi induk) | note di footer slide 6 | label "DER" di sheet `WP` (0,249) sebenarnya utang/aset |
| 5 | ROCE slide 5 = definisi CFI: EBIT ÷ (total aset − liabilitas jangka pendek), 2,1% vs 8,0%. Sheet `Ratios` memakai definisi lain: 2,3% vs 10,1% | note di footer slide 5 | pilih satu |

## 2. Koreksi terhadap catatan kuning di v1

- **"COGS naik karena pajak karbon"** → **tidak dipakai.** Pajak karbon domestik belum pernah dipungut (UU HPP; rencana 1 Apr 2022 tertunda; DDTC, 12 Okt 2026). Di slide 13 ditulis sebagai *risiko ke depan*. Penggerak COGS yang terbukti: overhead pabrik (+9,0 pp) dan energi (+4,0 pp).
- **"Jawa pusat produksi, defisit di Sumatra/Timur/Bali-Nusra"** → **didrop.** Tidak ditemukan data kapasitas vs konsumsi per wilayah. Yang dipakai: Jawa ±51% pasar 2025 dan Kalimantan −16,1% (INTP public expose 2026, sumber Kemenperin/BPS).
- **"Utilisasi ~67% flat"** → 67% itu **utilisasi SIG**. Utilisasi total industri (domestik + ekspor) 61–65%; slide 13 memakai angka industri.
- **"Ekspor klinker"** → pendapatan klinker relatif datar sepanjang FY21–25 (3,2 → 3,4 T), tapi naik 26% di FY25. Volume regional/ekspor ±8 Mt (21% volume). Ditulis apa adanya.
- **Eco-brand 46% pasar** (catatan di sheet `Ratios`, sumber earnings call FY25) → **tidak dipakai** karena belum bisa diverifikasi ke dokumennya.

## 3. Log sumber klaim slide 13 (data non-Excel)

| Klaim | Status | Sumber |
|---|---|---|
| Kapasitas INTP 33,5 Mt; pangsa domestik INTP 29,1% (2025) | FACT | Siaran pers kinerja FY2025 Indocement |
| Jawa ±51% pasar 2025; Kalimantan −16,1% | FACT-BOUNDED (dibaca dari slide peta) | INTP Paparan Publik 2026, hlm. 8 (Kemenperin, BPS) |
| 3 inisiatif strategis SIG sejak Jul 2025; market share 46,5% → 46,8% (1H25 → 9M25); harga fighting brand +2,2% (Jun–Sep 2025) | FACT (klaim perusahaan) | SIG Earnings Call 9M25, 5 Nov 2025, hlm. 6 & 13 |
| Perampingan entitas 40 → 12, eksekusi Q2 2026 | FACT-BOUNDED (media sekuritas) | BCA Sekuritas/IQPlus, 21 Mei 2026 |
| Pajak karbon belum berjalan | FACT | DDTC News, 12 Okt 2026 |
| HBA USD 121 → 277/t | FACT | sheet `Business Context` (ESDM) |

**Belum dipakai, perlu cek sendiri kalau mau masuk:**
- Kerja sama ekspor dengan Taiheiyo dari Tuban (0,5–1 Mt/thn, pertengahan 2026).
- Laba 1H26 +409% (opini kolom ANTARA).

## 4. Catatan desain

- **Warna:** SIG navy `#12274C`, INTP merah `#C00000`. Karena merah dipakai INTP, pemburukan ditandai ▼ dan angka FY25 bold, bukan dengan warna.
- **Validator palet (skill dataviz):**
  - Pemisahan navy vs merah **lolos** (ΔE CVD 18, normal 35).
  - Cek "lightness band/chroma" FAIL karena navy brand lebih gelap dari band palet generik. Diterima karena warnanya wajib dari template.
  - Tint terang (`#CBEBF8`, `#69C6F0`) kontrasnya rendah, dikompensasi data label.
- **Chart baru** (2 di slide 5, 3 di slide 6, 1 di slide 13) datanya **ter-embed**. Chart kiri slide 5 masih ter-link ke Excel OneDrive (warna diganti, data tidak diubah).
- **QA:** validator OOXML lolos (`--original` v1). Diff kanonik: hanya slide 5, 6, 13 dan chart slide 5 yang berubah isi. 75 slide utuh. Render diperiksa lewat PowerPoint.

## 5. Slide 6 v3: Leverage & Liquidity (Efficiency dihapus)

**File:** `slides\Slide_06_Leverage_Liquidity_v3.pptx` (1 slide, 0,26 MB). Copy lalu paste ke deck, menggantikan slide 6 lama. Build: `python scripts\build_slide6_v3.py`.

**Definisi:**
- Utang berbunga = pinjaman bank + obligasi + sewa + dana syirkah temporer.
- Debt-to-equity dan Debt ÷ net profit memakai ekuitas dan laba **porsi pemilik entitas induk**.
- Debt ÷ net profit dihitung dari sel (utang berbunga ÷ laba ke induk); INTP memakai laba inti tanpa one-off.

**Nuansa yang sengaja tidak di slide** (ada di speaker notes):
- Repayability berbasis kas tidak memburuk. CFO ÷ utang 34% → 36%, EBITDA ÷ bunga 4,8x → 5,4x. Karena itu wording slide memakai *"did not improve"*, bukan *"eroded"*. Yang benar-benar memburuk adalah basis laba: EBIT ÷ bunga 2,9x → 1,5x dan utang ÷ laba bersih 10,0x → 54,5x.
- Saldo laba hanya naik Rp32,9 T → Rp34,4 T, dan turun di FY25 karena dividen (Rp0,65 T) melebihi laba (Rp0,19 T).
- Watch item: utang jatuh tempo ≤1 tahun naik dari Rp2,2 T (FY24) ke Rp4,2 T (FY25), sementara kas Rp4,6 T.

**Sumber non-Excel:** rating PEFINDO idAA+ (Agu 2022/2023) → idAAA (Agu 2024), reaffirm idAAA/Stable Agu 2025; serta kutipan *"continued deleveraging efforts led to a reduction in finance costs"*. Keduanya dari SIG Earnings Call 9M25 (5 Nov 2025), hlm. 10–11. Status: FACT-BOUNDED (klaim perusahaan yang mengutip PEFINDO).

**Follow-up:** Efficiency masih muncul di Executive Summary (slide 1), quick scan (slide 19), dan Key Ratios (slide 26). Selaraskan, atau pindahkan ke RCA profitabilitas sebagai dampak kapasitas idle (fixed asset turnover 0,65x vs INTP 0,94x).

## 6. Slide 9: Revenue = Price x Volume (revisi v3)

**File:** `slides\Slide_09_Revenue_Price_Volume_v3.pptx` (1 slide, 0,26 MB). Versi sebelumnya (`..._v2.pptx` dan versi pertama) dibiarkan. Build: `python scripts\build_slide9_revenue.py`. Script berisi assert untuk setiap klaim di slide, plus guard "tidak ada titik koma" di slide maupun di speaker notes.

**v3 (18 Okt 2026, setelah pertanyaan user "kenapa nggak pakai index lagi" dan "PPI sudah di-take into account belum"):**
- Chart bawah kembali ke **indeks FY21 = 100**. Pembandingnya digambar sebagai garis putus-putus abu: PPI sektor di chart harga, permintaan nasional di chart volume. Garis sumbu ada di 100 sebagai acuan level FY21.
- Di v2 garisnya nominal. PPI hanya muncul di kotak dan headline, dan angka riil hanya ada di speaker notes. Sekarang **angka riil masuk ke slide**:
  - headline: "fell 4% in real terms as sector producer prices rose 12%";
  - KTW 1: "Both fell about 4% in real terms";
  - footer: "Real = deflated by sector PPI".
- **Label per titik dihapus.** Tiga garis indeks harga berimpit dalam selisih <1–3 poin, jadi label akan menumpuk dan menutupi garis PPI. Angka ujungnya ada di kotak "Change 21–25".

**Yang berubah dari versi pertama:**
- **Sumber harga dan volume:** `IS (Presentation)` → blok "B. Per ton" (baris Net revenue dan Volume); SIG kolom C–G, INTP kolom J–N. Angkanya sama dengan sebelumnya (SIG 865 → 929, INTP 822 → 889). Script mengecek harga × volume = net revenue, dan net revenue = `IS (Build)`.
- **Deflator:** CPI diganti PPI BPS Industri Barang Galian Bukan Logam (KBLI 23).
- **Chart bawah** mengikuti chart referensi user:
  - nilai aktual, bukan indeks;
  - SIG `#083358`, INTP `#358FD1`, garis 1,5 pt;
  - label kotak di setiap titik, legend di kanan, kotak "Change 21–25".
  - Garis CPI/pasar dan baris catatan di bawah chart dihapus.
- **Wording:**
  - "an industry-wide cap" diganti "both trailed sector producer prices";
  - semua titik koma dihapus;
  - footer placeholder dilebarkan sampai 790 pt supaya catatan muat 2 baris.

**Definisi dan caveat:**
- **Harga = net revenue ÷ total volume** (semen + klinker, termasuk regional/ekspor), sama untuk SIG dan INTP.
  - Pembilang SIG ikut memuat revenue non-semen, jadi sebagian dari +7,4% berasal dari lini non-semen.
  - Revenue semen + klinker per ton saja hanya +5,4% nominal.
- **PPI KBLI 23, Q4 ke Q4 (Tabel 4 tiap BRS IHP), chain-link antar-basis:**
  - 2022: +5,01%. BRS No. 14/02/Th. XXVI, 1 Feb 2023, "Industri Barang Mineral Bukan Logam", 2010=100, 156,13 → 163,95 ([rilis](https://www.bps.go.id/id/pressrelease/2023/02/01/2017/harga-produsen-mengalami-inflasi-0-90-persen-di-triwulan-iv-2022.html)).
  - 2023: +2,93% dan 2024: +1,91%. BRS IHP Triwulan IV-2024, 3 Feb 2025, 2016=100 ([rilis](https://www.bps.go.id/id/pressrelease/2025/02/03/2402/harga-produsen-mengalami-inflasi-0-59-persen-di-triwulan-iv-2024-.html)).
  - 2025: +1,85%. BRS No. 15/02/Th. XXIX, 2 Feb 2026, 117,87 → 120,05 ([rilis](https://www.bps.go.id/id/pressrelease/2026/02/02/2542/harga-produsen-mengalami-inflasi-0-76-persen-di-triwulan-iv-2025.html)).
  - Kumulatif +12,2%. SIG relatif −4,3%, INTP −3,6%.
- **Batasan PPI:**
  - Indeks memuat semen SIG/INTP sendiri, plus kaca, keramik, dan produk beton. Jadi "real" di slide berarti "dideflasi dengan PPI sektor" (didefinisikan di footer). Ini benchmark harga sektor, bukan deflator konsumen.
  - Harga per ton adalah rata-rata tahunan dan ikut terpengaruh mix (klinker, ekspor, fighting brand).
- **Robustness:**
  - CPI +13,2% → SIG −5,1%.
  - IHPB bahan bangunan +10,4% → SIG −2,8%.
  - Arahnya sama di semua deflator (di-assert di script).
- **Pola harga:** puncak di FY22, lalu turun 3 tahun berturut-turut (SIG −5,7%, INTP −4,2%), sementara PPI naik +6,8%.
- **Untuk slide COGS:** biaya per ton SIG +20% vs harga +7%; INTP +9% vs +8%.
- **Volume SIG cocok dengan sumber resmi:**
  - FY24 → FY25 38,27 → 37,93 Mt: [Earnings Call FY25](https://sig.id/storage/downloads/presentasi-corp/id/2025-presentasi-korp/earnings-call-fy-25-ppt-smgr.pdf), hlm. 6.
  - FY22 → FY23 36.921 → 40.621 rb ton: [Corp Presentation Mar 2024](https://sig.id/storage/downloads/presentasi-corp/id/2024-presentasi-korp/sig-corporate-presentation-march-2024.pdf), hlm. PDF 21.
  - FY21 42,4 Mt: AR2022, sudah termasuk SMBR. Sesuai dengan [Cemnet](https://www.cemnet.com/News/story/174575/semen-indonesia-sales-volume-declines-3-in-2022.html) "36,92 Mt, down 12,9%".
- **Permintaan nasional** (ASI/Asperssi via MD&A AR, `IS (Build)` baris 77): 66,2 → 63,91 Mt (−3,5%).
  - FY24 → FY25: Earnings Call FY25, hlm. 7.
  - FY23 65,5 (+3,5%): Corp Presentation Mar 2024, hlm. PDF 6 dan 8.
  - FY22 63,0 (−4,8%, artinya FY21 66,2): [Earnings Call FY2022](https://www.sig.id/storage/downloads/presentasi-corp/id/2022-presentasi-korp/earningscall-fy2022.pdf).
  - ASI kemudian merevisi FY22 ke ±63,3 (dipakai di sheet `Business Context`). Ini tidak memengaruhi slide, karena slide hanya memakai FY21 → FY25.
  - Seri Kemenperin ([INTP Paparan Publik 2026](https://www.indocement.co.id/media/post/materi-paparan-publik-2026-exclude-fs-ina-8de8c0dd02a64295547.pdf), hlm. 5) memberi 63,9 → 62,4 (−2,3%). Kesimpulan "pasar turun ±3%" tetap berlaku.
- **Volume INTP +11% sebagian inorganik:** volume Semen Bosowa yang dijual INTP dan Semen Grobogan. Ini disebut di caption chart.
  - Catatan workbook tidak konsisten soal tahun Grobogan (FY23 di `INTP_3FS`, FY24 di `Business Context`), jadi slide tidak menyebut tahun.
- **Dekomposisi volume domestik:** −4,1 Mt = efek pasar −1,2 Mt + efek pangsa −2,9 Mt (71%).
- **"Regional & export"** termasuk penjualan TLCC di Vietnam, jadi bukan murni ekspor dari Indonesia.
- **Beban tetap** = tenaga kerja + D&A + sewa + overhead pabrik (di COGS): Rp336 rb → Rp436 rb per ton (+30%).
  - D&A per ton hampir datar, jadi kenaikannya terutama di overhead.
- **"~35% idle capacity"** = 1 − (permintaan domestik 63,9 + ekspor nasional 13,7) ÷ kapasitas 119,9 Mt. Sumber: sheet `Business Context`, FY25; dihitung di script.
- **Klaim nilai tambah klinker lebih rendah belum terkuantifikasi**, karena volume klinker tidak ada di workbook.
- **Validator palet:**
  - Pemisahan SIG/INTP lolos (CVD ΔE 31, kontras ≥ 3:1).
  - Cek "lightness band/chroma" FAIL karena navy `#083358` lebih gelap dari band generik. Diterima karena warnanya mengikuti referensi user.

## 7. Follow-up untuk slide lain

- **Warna:** slide 9 memakai warna referensi user (SIG `#083358`, INTP `#358FD1`), sedangkan slide 5 dan 6 masih SIG navy `#12274C` / INTP merah `#C00000`. Selaraskan bila diminta.
- **`sig_data.ratio()`** sekarang membaca sheet `Ratios` (rename dari `Ratios v2`, struktur sama). Cek nomor kolom "No" sebelum build ulang slide 6.

- **Slide 4 & 12:** pie memakai SIG merah / INTP navy, jadi sekarang **terbalik** dari skema baru. Perlu disamakan.
- **Slide 12:** isinya tumpang tindih dengan slide 13 (oversupply, HBA, market share) dan masih ada chart dobel. Kandidat dihapus atau dijadikan backup.
- **Slide 4:** headline masih slang ("gets effed"); kotak kuning "Drop n rise by brp PP" bisa diisi: SIG −4,6 pp, INTP +3,7 pp, Others +0,9 pp.
- **Slide 9:** tracker masih "Financial Analysis" saja, sedangkan slide 10 memakai "Financial Analysis – Profitability Analysis". Samakan kalau perlu.

## 8. Slide 10: OPEX deep dive

**File:** `slides\Slide_10_Opex_Deep_Dive_v2.pptx` (1 slide, 0,26 MB). Versi 1 (`Slide_10_Opex_Deep_Dive.pptx`) dibiarkan. Build: `python scripts\build_slide10_opex.py`. Script berisi assert untuk setiap klaim dan guard tanpa titik koma.

**Aturan sumber (user, 18 Okt 2026):**
- Semua angka opex **hanya dari `IS (Presentation)`**, tidak ada lagi `IS (Build)` atau `INTP_3FS`.
- Penjualan excl. angkut = "SG&A excl. freight" (baris 63) − G&A reported (baris 17).
- Akibatnya temuan penurunan nilai piutang FY25 (sumbernya `IS (Build)`) dihapus dari slide dan notes.

**Definisi:**
- **Opex inti** = SG&A excl. freight (basis comparable). Definisinya sama untuk SIG dan INTP.
- **Other operating income (expense)** neto ditampilkan terpisah sebagai baris pendukung di bawah chart kiri. Bukan bagian opex inti.
- **Pembanding** = PPI BPS Barang Galian Bukan Logam (KBLI 23), sama dengan slide 9. Konstantanya sekarang di `sig_data.py` (`PPI_KBLI23_Q4_YOY` + `cumulative()`) dan dipakai slide 9 dan 10. CPI hanya di notes sebagai cek.

**Verdict atas saran user:**
- **IS (Presentation) saja:** setuju.
- **Masukkan other operating ke opex:** sebagian. Ditampilkan terpisah karena barisnya neto dan tandanya bolak-balik. Kalau dimasukkan, CAGR SIG 2,9% (sama dengan PPI 2,9%) dan INTP 12,5%, jadi kesimpulannya tetap. Nilai FY25 −Rp0,29 T setara 23% EBIT.
- **2 chart + kotak KTW memanjang, CAGR di chart kiri:** setuju, dipakai.
- **Stacked "Selling + G&A":** setuju, asalkan penjualan excl. angkut. Beban penjualan reported turun 38%, tapi 97% dari penurunan itu ongkos angkut yang pindah ke COGS (2,37 → 1,17 T).
- **"Jangan pakai adjusted INTP kalau SIG reported":** prinsip konsistensinya setuju, arahnya tidak. Basis comparable dipakai untuk keduanya. Secara reported hasilnya terbalik (SIG 15,7% vs INTP 20,8%); pengecekan ini memakai `INTP_3FS`, di luar slide.
- **PPI, bukan CPI:** setuju, dengan framing "harga produsen sektor". CPI 3,1% vs PPI 2,9% memberi hasil yang sama.

**Angka kunci (IS (Presentation)):**
- **Opex SIG** Rp4,12 T → Rp4,36 T (CAGR +1,5%): penjualan excl. angkut −1,2%, G&A +2,2%. INTP Rp1,11 T → Rp1,47 T (+7,2%).
- **Rasio opex:** SIG 11,2% → 12,4%, INTP 7,5% → 8,3%. Gap 3,7 → 4,1 pp. Avg 21–25: 11,4% vs 7,7%.
- **Other operating neto:**
  - SIG: −0,03 / +0,06 / +0,19 / −0,13 / −0,29 T.
  - INTP: +0,16 / +0,42 / +0,06 / +0,06 / −0,05 T.
- **Bridge EBIT** −10,6 pp = gross margin −8,7 + opex −1,2 + operasi lain −0,7.
- **Opex per ton FY25:** SIG Rp115 rb, INTP Rp74 rb.

**Framing gap:**
- Wording slide: "structural gap that barely moved". Wireframe menulis "masih wajar", padahal secara level tidak.
- Penjelasannya (holding dengan banyak entitas, multi-brand) berstatus INTERPRETATION. Pendukungnya FACT-BOUNDED: keterbukaan SIG soal likuidasi entitas Mei–Jul 2026, dan corporate presentation SIG (multi-brand).

## 9. Slide 13: Liquidity

### v2/v3 (versi BOD, berlaku)
**File:** `slides\Slide_13_Liquidity_v3.pptx` (0,25 MB). Build: `python scripts\build_slide13_liquidity.py` (sekarang menulis v3).

**Revisi v3:**
- Satuan diganti dari "per Rp100 aset lancar" menjadi **% dari aset lancar** (91% vs 54%, kas 27% vs 52%), sesuai permintaan user.
- Disimpan sebagai file v3 karena v2 sedang terbuka di PowerPoint user.

**Feedback user yang diterapkan:**
- Tanpa tag Now/Next, tanpa chip kelayakan, tanpa catatan meta.
- Kolom "Recommendation | Supporting evidence".
- Tanpa kolom dampak (estimasi hanya di notes).
- Footer cukup satu baris sumber.
- KTW berpola "against what SIG holds / generates".
- Pengulangan dibuang.

**Koreksi analisis:**
- Versi 1 membandingkan SIG adjusted 1,74x dengan INTP reported 1,84x, jadi tidak apple-to-apple. Kalau INTP di-adjust dengan cara yang sama, hasilnya 2,45x.
- Driver sebenarnya (% dari aset lancar FY25):
  - utang dagang SIG 45% vs INTP 14%;
  - utang berbunga jangka pendek 24 vs 13;
  - uang muka 9 vs 1;
  - lainnya 13 vs 26;
  - total 91 vs 54;
  - kas 27% vs 52%.

**Rekomendasi:** hanya aksi yang (a) bisa dieksekusi SIG atau punya jalur persetujuan formal, (b) pernah dilakukan SIG, dan (c) efek kasnya terasa dalam ±12 bulan.
1. **Refinance utang jangka pendek** ke obligasi atau pinjaman 3–5 tahun, plus standby line committed. Bukti:
   - idAAA stable (Agu 2026);
   - Rp2,3 T dilunasi Jan–Jul 2026;
   - Rp0,9 T di-refinance dengan bunga 1 pp lebih rendah setelah integrasi SMBR (SIG Corporate Presentation Mar 2024, hlm. 19).
2. **Kebijakan dividen terkait kecukupan kas**, diusulkan ke Danantara. Bukti: payout FY2023 hanya 26%; Danantara menetapkan dividen per kebutuhan holding.
3. **Percepat penagihan** lewat deposit/cash-before-delivery dan tempo piutang lebih pendek, tanpa memperpanjang tempo supplier. Bukti: uang muka pelanggan naik dari Rp0,1 T ke Rp1,6 T.

**Ditolak, alasannya ada di speaker notes:** cash pooling (dampak likuiditas kecil), scrip, asset recycling, sekuritisasi, shareholder loan DAM, dan reverse factoring.

**Estimasi dampak, hanya di notes:**
- current ratio ±1,50x setelah refinance (1,46x tanpa sewa);
- ±Rp1,1 T per tahun dari payout 50% di level laba FY22–23;
- ±Rp0,5 T per 5 hari piutang.

### v1 (digantikan)
**File:** `slides\Slide_13_Liquidity.pptx` (1 slide, 0,25 MB). Build: `python scripts\build_slide13_liquidity.py`. Script berisi assert untuk setiap angka, komposisi liabilitas lancar = total, dan guard tanpa titik koma.

**Layout** (mengikuti wireframe):
- **Kiri:** tabel framing (SIG FY21, FY25, CAGR, INTP FY25) dan panel KTW "where the risk sits".
- **Kanan atas:** current ratio reported dan adjusted, ditambah tabel liabilitas lancar.
- **Kanan bawah:** tabel 5 tuas (gaya P10) dengan chip kelayakan, preseden, dan baris "Avoid".

**Koreksi saat build** (riset di bawah sudah disesuaikan):
- Quick ratio FY21 = **0,77x**, bukan 0,78x.
- Hari utang dagang INTP dihitung dengan basis yang sama dengan SIG (delivered cost termasuk angkut): **41 hari**, bukan 49. SIG 98 hari.
- Quick ratio sempat 1,0x di FY22. Karena itu KTW berbunyi "buffers hold, but thinner than in FY22", bukan "steady".

**Sumber data:**
- SIG dari `BS (Presentation)`, `CF (Presentation)`, `IS (Presentation)`.
- INTP dari `INTP_3FS`, karena neraca INTP hanya ada di sana.
- Fakta non-workbook (payout, kebijakan Danantara, rating, preseden) disimpan sebagai konstanta di script, dengan sumber di speaker notes.

### Riset pendukung

**Framing data (workbook):**
- Cash ÷ utang berbunga ≤1 thn: 0,74x → 1,09x (CAGR +10,2%), INTP 4,13x.
- Quick ratio (aset lancar − persediaan) ÷ liabilitas lancar: 0,77x → 0,79x (1,0x di FY22), INTP 1,48x.
- CFO ÷ liabilitas lancar: 0,47x → 0,24x (−15,9%/thn), INTP 0,60x.
- Current ratio: 1,11x → 1,10x, INTP 1,84x.
  - FY25 tanpa uang muka pelanggan: 1,22x.
  - FY25 tanpa uang muka dan utang ≤1 thn: 1,74x.
- **Liabilitas lancar FY25 Rp15,8 T:**
  - utang usaha Rp7,8 T (98 hari, INTP 41, basis delivered cost);
  - utang berbunga ≤1 thn Rp4,2 T (hampir 2x FY24);
  - uang muka pelanggan Rp1,6 T (FY21 Rp0,1 T).
- **Sources/uses 12 bulan:** 1,49x, turun ke 1,29x kalau CFO −30%.
- **Penyebab:** CFO turun dari Rp6,9 T ke Rp3,7 T, sejalan dengan EBITDA (Rp8,6 T → Rp4,5 T).

**Dividen dan Danantara:**
- **Riwayat payout:** 26% (TB2023, Rp572 M) → 90% (TB2024, Rp648,7 M, [Bisnis](https://market.bisnis.com/read/20250523/192/1879565/tok-smgr-tebar-dividen-rp64875-miliar-buat-pemegang-saham)) → 100% (TB2025, Rp190,8 M, RUPST 8 Mei 2026, [Bareksa](https://www.bareksa.com/berita/saham/2026-05-11/sig-smgr-bagikan-100-laba-2025-sebagai-dividen-ini-potensi-yieldnya)). DAM memegang 50,69%.
- **Kebijakan dividen:** COO Danantara menyatakan dividen tidak lagi fixed dan mengikuti RKAP tiap holding (Feb 2026, [Bisnis](https://market.bisnis.com/read/20260218/192/1953762/dividen-bumn-2026-intip-kisi-kisi-danantara)).
- **Tidak ke APBN:** sesuai UU 1/2025, dividen tidak disetor ke APBN tetapi diinvestasikan, termasuk untuk menyelamatkan BUMN bermasalah (30 Sep 2026, [RCTI+/Okezone](https://www.rctiplus.com/news/detail/ekonomi/5484048/dony-oskaria--danantara-tak-ada-rencana-setor-dividen-rp120-triliun-ke-negara)).
- **Konsentrasi:** 97% dividen BUMN berasal dari 15 perusahaan ([CNBC](https://www.cnbcindonesia.com/market/20261006141738-19-773795/danantara-soal-kinerja-bumn-97-dividen-hanya-dari-15-perusahaan)).

**Refinancing:**
- PEFINDO idAAA stable ditegaskan 11 Agu 2026 ([PEFINDO](https://www.pefindo.com/rating-action-reports/rating-report/SMGR)).
- Kewajiban turun Rp2,3 T sepanjang Jan–Jul 2026 ([Stabilitas](https://www.stabilitas.id/pangkas-kewajiban-rp23-triliun-sig-pertahankan-peringkat-idaaa-dari-pefindo/)).
- Obligasi Rp714 M jatuh tempo 28 Mei 2026, dilunasi dengan dana internal ([BCA Sekuritas](https://bcasekuritas.co.id/latest-news/news/pefindo-berikan-peringkat-idaaa-untuk-obligasi-semen-indonesia)).
- PSAK 201 / IAS 1 par. 73: utang dengan hak rollover ≥12 bulan diklasifikasikan non-lancar.

**Benchmark akrobat finansial (preseden):**
- **Scrip dividend:** Repsol 2012–2020; 63,6% pemegang saham memilih saham ([Repsol](https://www.repsol.com/content/dam/repsol-corporate/en_gb/accionistas-e-inversores/pdfs/Scrip_dividend_OTCQX_tcm14-13257.pdf)).
- **Dukungan pemegang saham:** Temasek menanggung (underwrite) raise SIA S$15 miliar (2020, [Temasek](https://www.temasek.com.sg/en/news-and-resources/stories/t50/50-by-fifty/partnering-for-resilience-positioning-sia-for-a-strong-comeback)). DAM memberi shareholder loan Rp6,65 T ke Garuda (Jun 2025, [Bisnis](https://market.bisnis.com/read/20250624/192/1887788/profil-danantara-asset-management-dam-yang-suntik-garuda-giaa-rp665-triliun)). Perpetual bond BUMN China dicatat sebagai ekuitas (Bloomberg 2017).
- **Asset recycling:**
  - Jasa Marga ke INA (2021, [Bisnis](https://market.bisnis.com/read/20210308/192/1365266/ini-9-ruas-tol-yang-ditawarkan-jasa-marga-jsmr-ke-swf-indonesia)).
  - Holcim melepas Holcim Indonesia ke SIG (2019, [Holcim](https://www.holcim.com/closes-divestment-holcim-indonesia-accelerates-deleveraging)).
  - Cemex Operation Resilience; S&P investment grade Mar 2024 ([Business Wire](https://www.businesswire.com/news/home/20240313595342/en/Standard-Poors-Upgrades-Cemex-to-Investment-Grade)).
- **Sekuritisasi (KIK-EBA):**
  - Indonesia Power (PLN) 2017, maks. Rp4 T ([Detik](https://finance.detik.com/bursa-dan-valas/d-3650771/produk-sekuritisasi-aset-anak-usaha-pln-kelebihan-permintaan-27-kali)).
  - Jasa Marga 2017, Rp2 T ([CNN Indonesia](https://www.cnnindonesia.com/ekonomi/20170831092433-92-238564/sekuritisasi-aset-jasa-marga-oversubscribed-jadi-rp5-triliun)).
  - Piutang Cemex US$755 jt, tetap dicatat sebagai utang ([SEC 20-F FY2024](https://www.sec.gov/Archives/edgar/data/1076378/000119312525097693/Financial_Report.xlsx)).
- **Cash pooling:**
  - SIG sudah punya notional pooling dengan BNI (2019, [Kontan](https://keuangan.kontan.co.id/news/bni-geber-bisnis-layanan-pengeloaan-kas)) dan pendanaan grup terpusat (2022, [Bisnis](https://market.bisnis.com/read/20221015/192/1587878/semen-indonesia-smgr-beri-pinjaman-ke-anak-usaha-di-aceh)).
  - Pembanding: MIND ID–Timah–BNI (Jul 2026).
- **Hindari:** reverse factoring untuk memperpanjang tempo supplier, seperti kasus Carillion 2018, yang oleh Fitch disebut "hidden debt loophole" ([Construction Index](https://www.theconstructionindex.co.uk/news/view/carillion-hid-debt-burden-of-its-reverse-factoring)).

**Masih terbuka:**
- Dividen saham perlu cek legal; tidak ditemukan preseden BEI 2025–26.
- Ambang angka S&P (sources/uses) belum terverifikasi dari sumber primer, jadi tidak dikutip di slide.

## 10. Slide 4: Financial Analysis – Overview (bagian kiri direvisi)

**File:** `slides\Slide_04_Overview_v2.pptx` (0,43 MB). Build: `python scripts\build_slide04_overview.py`.

**Basis slide:**
- Slide 4 diambil dari deck kerja user di **OneDrive** ("SIG Finance Test/Financial Analysis SIG (SMGR) v1.pptx"). File v1/v2 di `D:\Download` sudah usang.
- Salinan lokal dibuat lewat PowerPoint `SaveCopyAs` ke scratchpad. Deck yang sedang terbuka tidak diubah.

**Yang berubah:** bagian kiri meniru pola bagian kiri slide 3, sesuai catatan reviewer @YG ("refer ke slide sebelumnya buat comparison metrik").
- **100% stacked:** revenue split (COGS / other costs / net profit), SIG vs INTP, FY21 vs FY25. Warnanya sama dengan chart user. Logo kecil diambil dari slide 3, dengan crop dipertahankan.
- **Kotak perubahan:** "COGS share of revenue, FY21 vs. FY25". SIG +8,8 pp dengan ▲ merah, INTP +0,4 pp dengan ▲ abu. Di bawahnya catatan "Driven by factory overhead and energy".
- **Compare bar FY25:**
  - revenue Rp35,2 T vs Rp17,7 T;
  - COGS Rp29,3 T vs Rp14,2 T;
  - net profit Rp0,19 T vs Rp1,58 T.
  
  Kalimat pembukanya: "Indocement earns 8x SIG's net profit on half the revenue".
- **Caption** diganti menjadi "Revenue split, % of revenue, FY21 vs. FY25".
- **Dihapus:** chart Rp T, panel teks pointer (isinya pindah ke speaker notes), dan catatan @YG.
- **Tidak disentuh:** judul (highlight kuning milik user), subheadline, chart NPM/ROA beserta Avg, dan Key Question.

**Data:** sheet `Profitability`, baris 32–35 (sumber chart lama).
- COGS = COGS + angkut.
- Net profit SIG = laba ke induk. INTP tanpa one-off.
- Perubahan pp dihitung dari angka yang ditampilkan (74,5 → 83,3), jadi tidak ada selisih pembulatan yang terlihat.
- Driver overhead dan energi dicek ke struktur biaya `IS (Build)`: overhead +9,0 pp, energi +4,0 pp dari revenue.
