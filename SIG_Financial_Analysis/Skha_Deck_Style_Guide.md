# Skha Deck Style Guide

Diekstrak dari `D:\Download\Financial Analysis SIG (SMGR) v1.pdf` (69 halaman) lewat parsing vektor PDF (PyMuPDF): warna fill/stroke, font, ukuran, dan koordinat persis, bukan tebakan visual.

Sudah **dikonfirmasi ulang** terhadap file sumbernya:
- `D:\Download\Financial Analysis SIG (SMGR) v1.pptx`: theme, master, layout, chart XML (lihat §10).
- `D:\Download\WP - Financial Analysis SIG (SMGR) v0.xlsx`: lihat §11.

Pembagian halaman:
- Hal. 1–30, 32 = deck SIG (masih WIP).
- Hal. 31, 33–69 = deck Pegadaian dengan template Skha yang sama. Ini yang dipakai sebagai **standar "final"**.
- Semua ukuran dalam **pt**, origin di kiri-atas. Skala PDF 1:1 dengan PowerPoint.

---

## 1. Kanvas & grid

**Kanvas** 960 × 540 pt (13,333 × 7,5 in, 16:9).

**Margin konten** kiri **40,5**, kanan **919,5**, sehingga lebar konten 879 pt. Teks judul mulai di x 47,9 (inset textbox ±7 pt).

| Zona | Posisi (pt) | Isi |
|---|---|---|
| Tracker | ikon diamond 15×15 di (40,5; 30,8); label mulai x 63, y 30–45 | breadcrumb section |
| Logo klien | x 848–894, y 17–44 | logo SIG (warna) |
| Stempel status | kotak 124,5×18 di (691; 32), outline merah 1 pt | "Preliminary" / "Need Confirmation" / "Illustrative" |
| Action title | x 47,9; baris 1 y 61,7, baris 2 y 83,5 (line ±21,8) | maks. 2 baris |
| Phase tracker (opsional) | y 99,8–113,2, rata kanan (±688–853) | chip progres "1a 1b 1c 2 3" |
| Body | y ±113 → ±508 | konten |
| Takeaway box | bagian bawah body, mis. y 411–493 atau 451–510 | so-what |
| Footer band | y 510,8–540, full width | source · logo skha · nomor halaman |

**Kotak placeholder resmi dari layout "6_Title Only"** (posisi x, y, lebar × tinggi):

| Placeholder | Posisi |
|---|---|
| Title (idx 0) | (40,6; 58,6) 878,7 × 55,5 |
| Tracker (idx 11) | (55,7; 26,7) 421 × 22,7 |
| Source (idx 10) | (40,6; 516,4) 510,2 × 19,8 |
| Slide number (idx 4) | (878,6; 519,3) |
| **Area body standar** (posisi kotak kuning di V1) | (40,6; 114,1) 878,7 × 379,9, berakhir di y ±494 |

**Sistem kolom yang terpakai:**

| Pola | Koordinat x | Gutter |
|---|---|---|
| 2 kolom setara | 40,5–513,7 · 539–919,5 | ±25 |
| 2 kolom asimetris (≈38/62) | 44–377 · 401–900 | ±24 |
| 3 kolom | 42–331 · 340–620 · 633–912 | ±10–13 |
| Main + sidebar | 41–718 · 728–919,5 (sidebar ±192) | ±10 |
| List + sidebar (hal. 38) | 44–673 · 691–919 | ±18 |

---

## 2. Tipografi

Keluarga font: **Univers**. Body memakai *Univers Light*, judul memakai *Univers Bold*.

Dikonfirmasi dari .pptx:
- Font theme: heading = `Univers`, body = `Univers Light`.
- Kedua font **ter-embed** di .pptx (Univers regular/bold/italic dan Univers Light), jadi aman walau laptop tidak punya Univers Bold.
- Typeface yang dipakai di XML: `Univers` (dengan b=1 untuk judul/tracker) dan `Univers Light`.
"Bold" di badan teks sebenarnya **Univers Light yang di-bold PowerPoint** (faux bold, terdeteksi dari render mode stroke di PDF). Judul memakai face Univers Bold asli.

| Peran | Font | Size | Warna | Catatan |
|---|---|---|---|---|
| Tracker | Univers Bold | 12 | `#1E64B0` | di kanan ikon diamond |
| Action title | Univers Bold | 18 | `#0F1F33` | sentence case, 1–2 baris, ±95–100 karakter/baris |
| Column headline | Univers Light **bold** | 12 | `#161616` | 1–2 baris, langsung diikuti rule biru |
| Exhibit caption | Univers Light | 10,5 | `#161616` | wajib ada satuan: "Revenue breakdown, Rp tn" |
| Body | Univers Light | 11,25–12 | `#161616` | sekunder `#363636` |
| Lead-in di bullet | Univers Light **bold** | 12 | `#161616` | "SIG cannot price above competitors: …" |
| Card heading | Univers Light | 14,25 | `#1A386E` | list kapabilitas (hal. 38) |
| Teks banner/bar | Univers Light | 11,25–12 | `#FFFFFF` | di atas `#12274C` / `#1E64B0` / `#368FCE` |
| Section tab (pill) | Univers Light *Italic* | 14,25 | `#FFFFFF` | hal. 3 |
| Angka KPI besar | Univers Light | 15,75 | `#C00000` memburuk · `#1E64B0` peer/baik · `#69C6F0` di bg gelap | |
| Label chart | Univers Light | 9–9,75 | `#161616`, atau putih di dalam bar gelap | |
| Body tabel | Univers Light | 10–10,9 | `#000000` | |
| Footnote | Univers Light | 6,75–8,25 | `#7F7F7F` | nomor superscript |
| Source | Univers Light | 8,25 | `#7F7F7F` | x 47,9, y ±523 |
| Nomor halaman | Univers Bold | 12 | `#0E1F31` | x ±896, y ±520 |
| Stempel status | Univers Light | 12 | `#FF0000` | |
| Emphasis di kalimat | font sama | | `#1E64B0` (atau `#12274C`) | penekanan pakai **warna**, bukan bold |
| Bullet | "•" Arial | 12 | `#000000` | bullet di x 48,3, teks di x 61,9 (hanging ±13,6) |

---

## 3. Palet

### 3a. Palet resmi template
Sumber: strip swatch di luar kanvas (x negatif) pada slide master. Ini yang paling otoritatif.

| Token | Hex | Dipakai untuk |
|---|---|---|
| Ink Navy | `#0F1F33` | judul, teks kunci, outline, bg slide gelap (`#0E1F32`) |
| Navy | `#12274C` | banner header, header tabel, rule di bawah header 3 kolom |
| Skha Blue | `#1E64B0` | tracker, rule kolom, aksen utama, circle bernomor, teks emphasis |
| Mid Blue | `#368FCE` | pill, banner sekunder, garis timeline putus-putus |
| Sky | `#69C6F0` | aksen tersier, step terang, highlight di bg gelap |
| Charcoal | `#595A5C` | netral gelap, state tidak aktif |
| Grey 1–4 | `#A6A6A6` · `#BFBFBF` · `#D9D9D9` · `#F2F2F2` | "others", inactive, panel netral |

**Theme color scheme "Skha Colors"** di .pptx. Pakai slot ini saat memilih warna dari theme:

| Slot | Hex |
|---|---|
| dk1 | `#0F1F33` |
| lt1 | `#FFFFFF` |
| dk2 | `#12274C` |
| lt2 | `#F0FAFE` |
| accent1 | `#1E64B0` |
| accent2 | `#368FCD` |
| accent3 | `#69C6F0` |
| accent4 | `#595959` |
| accent5 | `#7F7F7F` |
| accent6 | `#D8D8D8` |
| hlink | `#FFC000` |
| folHlink | `#3ECCB4` |

Strip swatch di master diberi label resmi:
- **Primary Colors**: `#0F1F33` `#12274C` `#1E64B0` `#368FCE` `#69C6F0`
- **Secondary Colors**: `#595A5C` `#DBDDDD` + ramp abu
- **For Highlighting Purposes**, masing-masing dari solid ke tint:

| Highlight | Solid → tint |
|---|---|
| Sky | `#69C6F0` · `#A0DCF1` · `#CBEBF8` · `#E0F3FC` |
| Green | `#9DC43B` · `#A8CF3B` · `#C3DB6C` · `#E7F1D4` |
| Yellow | `#FBCD0A` · `#FCD432` · `#FEDE55` · `#FFEFAC` |
| Red | `#EF4041` · `#F26B65` · `#F48889` · `#FAC6CA` |

Untuk sorotan positif/netral/negatif, ambil dari palet highlight ini, bukan dari merah/hijau default Office.

### 3b. Tint & warna pendukung (terpakai di deck)

| Hex | Peran |
|---|---|
| `#1A386E` | Deep navy: icon tile, card heading, kartu key question |
| `#F0FAFE` | panel paling terang |
| `#E1F4FC` | takeaway box, chip/tile |
| `#C3E8F9` · `#A5DDF6` | label band, tag chip |
| `#D6EFFA` | separator putus-putus |
| `#CCDCF0` | baseline chart dan border plot |
| `#E5E5E5` (0,5 pt) · `#C2C2C2` (0,75 pt) | hairline baris · separator grup |
| `#133E6A` · `#335B8E` | chevron value chain (gelap) |

### 3c. Semantik

| Peran | Hex |
|---|---|
| RAG kuat (legend chip) | Low `#C00000` · Med `#FFC000` · High `#00B050` |
| Heatmap soft | Low `#F7C3B3` · Med `#F6F8CC` · High `#D6EFBF` |
| Constraint/risiko | `#F48889` · `#FAC6CA` |
| Kemudahan/positif | `#C2EBFB` · `#F0FAFE` |
| Stempel | `#FF0000` |
| Merah brand SIG | `#F5333F` (**hanya di logo**) |

### 3d. Konvensi kuning (WIP)
Kotak placeholder/catatan = fill **`#FFC000`** + border 1 pt **`#0F1F33`**, teks 12–14 pt `#0F1F33`. Wajib hilang di versi final.

Jangan dikacaukan dengan pemakaian kuning yang sah: badge angka kecil kuning dan outline kuning untuk grup aktif di slide gelap. Bedanya, placeholder berupa kotak **besar berisi fill**, sedangkan aksen sah berupa badge kecil atau outline.

---

## 4. Chrome per tipe slide

Layout di .pptx: hanya 1 slide master, dengan 12 layout.

| Layout | Dipakai untuk |
|---|---|
| `6_Title Only` | content slide, semua slide SIG |
| `2_Title Only` | slide gelap |
| `Cover Page` | cover |
| `1_Title Only` | disclaimer / sub-tracker |
| `Tracker` | divider section |
| `Thank You` | penutup |
| `Appendix` | divider appendix |
| `End Document` | akhir dokumen |

**Content slide (terang)**
- Background putih.
- Ikon diamond dengan gradient linear 135°: `#0F1F33` (0%) → `#368FCD` (70%).
- Label tracker, logo SIG di kanan atas, action title.
- Footer band gradient horizontal putih → `#F0FAFE`, mulai y 510,8.
- Source di kiri, logo skha di x 816,8–875,2 / y 514,5–536,2, nomor halaman di kanan.

**Slide gelap** (agenda, key questions, guiding principles)
- Background `#0E1F32` dengan line-art gelombang tipis di sisi kanan.
- Judul putih bold 18, tracker putih Light 12, logo skha putih.
- Angka aksen `#69C6F0`.
- Kartu `#1A386E` / `#1E64B0` dengan opacity 90% dan border `#0F1F33`.
- Grup aktif diberi outline `#FFC000` 1 pt; circle angka kuning berisi angka 20 pt `#161616`.
- Kartu yang tidak aktif dibuat abu/redup.

**Cover**
- Foto konstruksi full-bleed dengan overlay navy.
- Logo skha putih di kiri atas (76,5; 59), logo SIG putih di kanan atas (766,5; 63).
- Judul putih bold ±27 pt dalam 2 baris, mulai x 83,4 / y ±313.
- Nama presenter bold 12,75, nama acara regular 12,75, dengan rule putih di bawahnya.

**Divider / closing**
- Background gelap, teks bold ±67 pt putih dengan aksen sky ("Thank **You**").

---

## 5. Pola layout (komponen)

| # | Pola | Spesifikasi |
|---|---|---|
| P1 | **Two-exhibit slide** (hal. 4–5) | Tiap kolom berurutan: headline 12 bold (1–2 baris) → rule **1,5 pt `#1E64B0`** selebar kolom → caption 10,5 (dengan satuan) → chart → legend. Di bawahnya **takeaway box `#E1F4FC`** full width: kalimat lead bold, lalu 2 bullet dengan lead-in bold. |
| P2 | **KPI callout panel** | Panel `#F2F2F2` (±145×246) di samping chart. Label 12; periode "FY21 → FY25" 9,75 `#404040`; angka besar 15,75 berwarna; delta "−5.1 pts" 9,75 `#404040`. |
| P3 | **Section tab** (hal. 3) | Pill gradient `#1E64B0`→`#69C6F0` dengan label putih italic 14,25, duduk di atas rule gradient full width. |
| P4 | **Label band + chip row** | Band `#A5DDF6` berisi label 12; di bawahnya chip `#E1F4FC` dengan jarak ±4,5. |
| P5 | **Value chain chevron** | Parallelogram dari gelap ke terang (`#133E6A` → `#335B8E` → `#368FCD`), label putih bold 12, deskripsi 11,25 di bawahnya. |
| P6 | **Timeline** | Garis putus `#368FCE` 1 pt [4 3]; node circle `#1E64B0` dengan stroke putih 1 pt; tahun 10,5 `#0F3258`; deskripsi 9,75 hitam. |
| P7 | **Tabel finansial** | Lihat §6. |
| P8 | **Main + sidebar KTW** | Sidebar `#F2F2F2` dengan border putus `#BFBFBF` 0,75. Header 14,25 bold `#0F1F33` plus ikon merah; item bernomor 12. |
| P9 | **List bernomor** (hal. 38) | Icon tile `#1A386E` 51×51 berisi ikon putih; heading 14,25 `#1A386E`; deskripsi 12 `#363636`; tag chip `#C3E8F9` di kanan; header section 15,75 bold `#69C6F0` dengan rule 4,5 pt `#69C6F0`. |
| P10 | **Tabel parameter/skor** (hal. 43) | Header tile `#12274C` (label) dan `#1E64B0` (benchmark); chip RAG di header; tile kategori `#1E64B0` dengan ikon; kotak definisi putih dengan putus-putus `#A6A6A6`; kolom parameter `#F0FAFE`; separator baris `#E5E5E5` 0,5 dan grup `#C2C2C2` 0,75; **tanpa garis vertikal**. |
| P11 | **Heatmap matrix** (hal. 58–59) | Sel memakai tint Low/Med/High; label baris berupa sel navy `#12274C` dengan teks putih; footnote 6,75 `#7F7F7F`. |
| P12 | **Process flow** (hal. 44, 49, 62) | Strip chevron header `#12274C` dengan teks putih; box outline `#1E64B0` 1 pt; diamond keputusan Yes/No. |
| P13 | **Funnel 3 langkah** (hal. 42) | Step dari terang ke gelap: `#69C6F0` → `#368FCE` → `#1E64B0`. Box di bawah tiap step: outline putus warna yang sama, fill `#F0FAFE`. |
| P14 | **Phase tracker chip** | Band `#F0FAFE` berisi chip; chip aktif `#12274C` dengan teks putih, chip tidak aktif abu terang. |
| P15 | **Gantt** (hal. 50, 64–66) | Band bulan `#1E64B0`; bar `#CBEBF8` / `#69C6F0`; kolom "We are here" outline kuning; milestone diamond `#F26B65` (critical) / `#A8CF3B` (key). |
| P16 | **Company profile** (hal. 54–57) | Panel kiri `#F2F2F2`: logo, deskripsi, bendera footprint, baris "Financial snapshot" (pill label `#368FCE` + bar nilai `#F2F2F2`). Di kanan, roadmap bentuk tangga. |

---

## 6. Tabel laporan keuangan (gaya V1, hal. 20–25)

**Warna baris**
- Header: `#002060`, teks putih 10–10,9. Sel kiri atas putih berisi "Income Statement (Rp juta)".
- Baris section ("Revenue", "COGS"): `#0070C0` dengan teks putih.
- Body: putih, teks hitam 10, tinggi baris ±13,1.
- Subtotal: `#DDEBF7`.
- Grand total: `#BDD7EE` dengan hairline atas `#002060`.

**Kolom**
- Label ±165, FY21–FY25 masing-masing ±76, CAGR (21-25) dan YoY (24-25) masing-masing ±51.
- Angka rata kanan.

**Format angka (Indonesia)**
- Ribuan pakai titik: `36.702.301`.
- Negatif dalam kurung: `(1.773.650)`.
- Persen pakai koma: `(1,0%)`.
- `n.m.` untuk tidak bermakna, `–` untuk nol.
- Baris memo (Free cash flow, Utang berbunga) ditulis *italic*.

**Footnote dan sumber**
- Footnote bernomor, lalu baris sumber: "Sumber: LK konsolidasian auditan SIG FY2021–2025 (Rp juta)."

> **Catatan:** biru-biru tabel ini adalah default Excel, mirip palet template tapi tidak identik. Opsi penyamaan tanpa mengubah struktur: `#002060`→`#12274C`, `#0070C0`→`#1E64B0`, `#DDEBF7`→`#E1F4FC`, `#BDD7EE`→`#C3E8F9`.

---

## 7. Chart

**Yang terlihat di V1**
- Slide 5 (paling bersih): tanpa border chart, tanpa gridline, tanpa garis sumbu Y; baseline `#CCDCF0` 0,75.
- Slide 4: masih ada border plot `#CCDCF0`.
- Label sumbu 9 pt, tapi warnanya tidak konsisten (`#3268AB` di slide 5, `#161616` di slide 4).
- Data label: putih 9 pt di dalam segmen gelap; total di atas bar memakai `#262626` / `#404040`.
- Legend di bawah, marker kotak 5×5, teks 9 pt. Judul di dalam chart 9,75 (mis. "Industry").
- Line chart: garis `#1F3864` 2,25 pt dengan marker bulat.
- Pie: separator slice putih 1,5 pt, label putih 9 pt di dalam slice.

**Inkonsistensi di V1**
- Warna campur aduk: default Office (`#1F4E79`, `#5B9BD5`, `#BFBFBF`, `#1F3864`, `#C00000`, `#A6A6A6`) bercampur ramp custom slide 5 (`#10719D`, `#6FA8DC`, `#C5DAEE`).
- Sebagian label chart masih memakai font **Aptos Narrow**.

**Spesifikasi chart terpadu yang diusulkan** (diterapkan saat memasukkan chart)

| Elemen | Spesifikasi |
|---|---|
| Font | Univers Light 9 pt; label sumbu `#595959`; data label putih di dalam fill gelap, `#262626` di luar |
| Entitas | **SIG `#12274C`** (fokus) · **INTP `#69C6F0`** · **Others `#BFBFBF`** |
| Merah | `#C00000` khusus untuk sorotan pemburukan (delta, KPI, callout) |
| Ramp stacked (gelap → terang) | `#12274C` · `#1E64B0` · `#368FCE` · `#69C6F0` · `#A5DDF6` · `#E1F4FC` |
| Kapasitas idle / unused | `#D9D9D9` |
| Garis | entitas fokus 2,25 pt `#12274C` dengan marker · peer 2,25 pt `#69C6F0` · benchmark/threshold putus-putus 1 pt `#7F7F7F` |
| Sumbu | baseline `#CCDCF0` 0,75; tanpa gridline; tanpa border; tanpa sumbu Y kalau sudah ada data label |
| Caption | selalu memuat metrik, satuan, periode: "Net profit margin, %, FY21–FY25" |

> V1 memakai **SIG = merah** (pie, KPI). Kalau tetap mau pakai merah brand SIG, pakai **konsisten satu skema** di seluruh deck, dan merah tidak lagi dipakai untuk "negatif".

---

## 8. Konvensi penulisan

- **Action title** berupa kalimat utuh yang menyatakan so-what. Maks. 2 baris pada 18 pt (±95–100 karakter/baris).
- **Column headline** berupa sub-klaim yang mendukung judul (1–2 baris). Caption menjelaskan *apa* yang diplot.
- **Bahasa:** deck Pegadaian memakai Bahasa Indonesia dengan istilah Inggris dalam *italic*; SIG V1 memakai English. Pilih satu per deck.
- **Bersihkan sebelum final:** slang ("effed", "WKWKWK", "sih", "boncos", "ingfo", "emg").
- **Format angka di V1 tercampur:**
  - Tabel memakai format Indonesia (`36.702.301`, `1,0%`).
  - Label chart memakai desimal titik (`36.7`).
  - Hal. 3 mencampur "56,5 MTPA" dengan "63,000+".
  - Samakan konvensinya. "Source:" vs "Sumber:" juga ikut bahasa deck.
- **Typo di V1:** Reccomendation, Profitabbility, divident, deleverating, repayabilty, ditompang (→ ditopang), DNA (→ D&A).

---

## 9. Status slide V1 (placeholder kuning)

| Hal. | Tracker / topik | Status |
|---|---|---|
| 1 | Executive Summary | kuning (draft poin) |
| 2 | Cover | final |
| 3 | About SIG | final |
| 4 | Competitive Landscape | 2 kuning ("Drop n rise by brp PP", "Key numbers benchmark against INTP") + slang di headline |
| 5 | Financial Analysis (revenue & margin vs INTP) | panel kanan "Margin and Returns" **kosong** |
| 6 | Quick scan leverage / liquidity / efficiency | 3 kuning |
| 7–9 | Problem tree / jawaban analisis | kuning |
| 10–11 | Conclusion / Recommendation | kuning |
| 12 | Market Landscape | chart dobel/bertumpuk (WIP) |
| 13 | Situation understanding | 3 kuning |
| 14–18 | Management response / liquidity / market response / what can SIG do | kosong |
| 19 | Performance overview (quick scan KPI) | 5 kuning |
| 20–25 | IS · CF ops · CF financing · BS aset · liabilitas · ekuitas | tabel final, **sidebar KTW kuning** |
| 26 | Financial Ratios | kuning |
| 27–29 | RCA Profitability · Debt Repayability · Short-term Liquidity | kuning |
| 30 | Recommendation | kuning |
| 31 | Disclaimer | **masih logo Pegadaian** |
| 32 | Appendix | kuning |

---

## 10. Catatan build

**Struktur file .pptx V1**
- 75 slide. Slide 70–75 hidden (draft Pegadaian).
- Satu master, 12 layout, font ter-embed.

**Cara drafting**
- Kerjakan langsung di .pptx V1 dengan layout `6_Title Only`.
- Isi placeholder: Title (idx 0), Tracker (idx 11), Source (idx 10).
- Footer gradient, diamond, logo skha, logo SIG, dan nomor halaman sudah otomatis dari layout. **Jangan digambar ulang.**

**Chart V1**
- 10 chart native di slide 4, 5, 12.
- Semuanya **ter-link ke Excel WP di OneDrive** (`Documents/SIG Finance Test/WP - Financial Analysis SIG (SMGR) v0.xlsx`).
- Kalau chart dibuat/diubah lewat script, datanya akan ter-embed dan link OneDrive putus. Konsekuensi ini perlu diputuskan dulu.
- Referensi gaya terbaik adalah **chart slide 5**: Univers Light 9, tanpa gridline, sumbu Y dihapus, gap 50, overlap 100. Chart slide 4/12 masih punya gridline, font theme/Aptos Narrow, dan warna default Office.

**Tabel laporan keuangan (slide 20–25)**
- Berupa **gambar WMF** hasil paste-special dari Excel (sheet `IS/BS/CF (Presentation)`).
- Slide 24 berupa **objek OLE Excel embedded**.
- Kalau mau tabel native yang bisa di-restyle, bangun ulang dari sheet Presentation.

**Speaker notes**
- Slide 8 berisi arahan: "Revenue ga drop (kasi info revenue mix + price increase) … Expenditure wise → vertical analysis".

**Font**
- Di laptop ini hanya `Univers LT Std 45 Light` yang terpasang, tapi tidak masalah selama kerja di .pptx asli karena font ter-embed.
- **Aset hasil ekstraksi** (`template_assets\`):
  - `tracker_diamond.png`, `sig_logo_color.png`, `skha_logo_color.png`
  - `skha_logo_white.png`, `skha_logo_white_cover.png`, `sig_logo_white_cover.png`
  - `bg_dark_section.png`, `bg_cover_photo.png`, `section_tab_pill_example.png`
- **Token mesin:** `style_tokens.json` (folder ini).

---

## 11. Peta workbook `WP - Financial Analysis SIG (SMGR) v0.xlsx`

| Sheet | Isi | Dipakai untuk |
|---|---|---|
| `Profitability` | issue tree profit (revenue/ton × volume, COGS/ton fixed vs variable), data chart revenue breakdown SIG vs INTP | chart slide 5, RCA profitability |
| `WP` | working paper per slide: revenue mix FY25, segmen, pemegang saham, subsidiaries + aset, volume & utilisasi, market share, EBITDA, DER, net debt/EBITDA | About, overview, quick scan |
| `Business Context` | story table FY21–25, demand vs kapasitas industri, market share tahunan + catatan, COGS/ton SIG vs INTP, HBA batubara (USD dan Rp/ton) | market landscape, competitive |
| `IS/BS/CF (Presentation)` | sumber tabel gambar di slide 20–25 | statement slides |
| `IS (Reported)`, `IS (Build)`, `BS`, `CF` | build-up dari LK auditan (COGS per sifat, ongkos angkut, NCI, D&A, EBITDA) | rekonsiliasi |
| `Vertical` | common-size IS (% pendapatan) dan BS (% aset) | vertical analysis COGS |
| `Ratios` | 45 rasio + ⚑ flag + catatan + 5 KTW | rasio |
| `Ratios v2` | 61 rasio SIG vs INTP dengan definisi CFI, **KTW per kelompok A–G**, **hasil uji hipotesis** (H1 profitabilitas TERBUKTI · H2 repayability = dampak turunan H1 · H3 likuiditas TIDAK TERBUKTI, watch item), **shortlist 12 rasio untuk slide** | quick scan + RCA |
| `INTP_3FS`, `Peer_INTP` | 3 statement INTP + turunan | benchmark |

**Inkonsistensi data yang perlu dikunci sebelum drafting**

1. **"COGS" di chart slide 5** sebenarnya **COGS + ongkos angkut & bongkar** (`Profitability` row 32 = IS (Build) row 21 + 25). Contoh FY21: 27.350 vs COGS IS 24.976.
   - Teks slide 5 "COGS rose from 75% to 83%" memakai definisi ini, padahal COGS murni 68,0% → 79,9%.
   - Rekomendasi: label menjadi "COGS + freight", atau ganti ke COGS murni.
2. **"Net Profit"** di chart dan KPI = **laba ke pemilik entitas induk** (2.047 → 191). Slide 20 memakai **laba konsolidasi** (2.117 → 179). Basisnya perlu disebut.
3. **Net margin INTP FY25:**
   - Slide 5 = **8,9%** (tanpa one-off divestasi Rp670 M).
   - `Ratios` = **12,7%** (termasuk one-off).
   - `Ratios v2` memakai versi tanpa one-off. Pakai satu versi saja.
4. **DER** punya 3 definisi:
   - `WP` "DER" 0,249 sebenarnya **utang berbunga / total aset**.
   - `Ratios`/`Ratios v2` D/E = utang berbunga (incl. syirkah) / **ekuitas induk**, 0,53 → 0,24.
5. **"Debt to EBITDA"** di `WP` (1,28) sebenarnya **net debt / EBITDA**. Shortlist `Ratios v2` memakai **IBD / EBITDA** (2,31).
6. **ROCE** punya 2 definisi:
   - `Ratios`: EBIT / (ekuitas + utang bersih) = 2,3% vs INTP 10,1%.
   - `Ratios v2` (definisi CFI): 2,1% vs INTP 8,0%.
7. **`WP` segmen** "Non-semen %" tertulis 1 (100%), padahal total segmen 45,1 T adalah angka **sebelum eliminasi**, jauh di atas revenue 35,2 T.
8. **Market share:** pie memakai 51,5% → 46,9% (dari `WP`). Slide 4 teks/visual dibulatkan 52% → 47%, sedangkan narasi menyebut "lost 4.5 mt (−11%)". Volume 42,4 → 37,9 = −4,5 mt (−10,6%): konsisten, cukup samakan pembulatannya.
