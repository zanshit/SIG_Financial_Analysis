"""Slide 13 v2 (BOD version): Liquidity.

Left: the wireframe's three liquidity metrics (SIG FY21, FY25, CAGR | INTP FY25) and the key takeaways.
Right: what holds SIG's current ratio down, measured like for like (current liabilities as % of current
assets, SIG vs INTP), and three recommendations with the supporting evidence for each.
Delivers ONLY this slide as a slim .pptx (template from v1, one layout, Univers fonts only).
SIG figures from BS/CF/IS (Presentation), parent profit from IS (Build) for the notes only, INTP from INTP_3FS.
Run:  python build_slide13_liquidity.py
"""
import sys
from pathlib import Path

from pptx.util import Pt

sys.path.insert(0, str(Path(__file__).parent))
from sig_data import block_row, check_numbers, row_values  # noqa: E402
from skha_helpers import (BODY, GREY, INK, INTP_BLUE, PANEL_GREY, SIG_NAVY, SKHA_BLUE, add_column_chart,  # noqa: E402
                          add_column_header, add_kpi_table, add_legend, add_line, add_rect, add_summary_boxes,
                          add_text, add_text_table, clear_slide, fill_text_frame, make_slim_deck, replace_runs_text)

SRC = Path(r"D:\Download\Financial Analysis SIG (SMGR) v1.pptx")
OUT = Path(r"D:\Download\SIG_Financial_Analysis\slides\Slide_13_Liquidity_v3.pptx")

# Outside the workbook (sources in the speaker notes):
PAYOUT_FY23 = 0.2636          # dividend on FY2023 profit, AGM 2024 (Bisnis, 23 May 2025)
DIV_FY25 = 190_840            # Rp juta, 100% of FY2025 profit, approved 8 May 2026 (Bareksa)

PAY_C, DEBT_C, ADV_C, OTH_C = SKHA_BLUE, INK, "69C6F0", "8C8C8C"   # validated: adjacent dE >= 16


def x(v, d=2):
    return f"{v:.{d}f}x"


def pct(v, d=1):
    """Signed percent with an en dash for negatives (no U+2212 dependency on Univers)."""
    return ("+" if v >= 0 else "\u2013") + f"{abs(v) * 100:.{d}f}%"


def t(v, d=1):
    return f"{v / 1e6:.{d}f}"


# ---- data ------------------------------------------------------------------------
B, C, I, P = "BS (Presentation)", "CF (Presentation)", "INTP_3FS", "IS (Presentation)"
cash, ar, inv, ca = (row_values(B, "Kas dan setara kas"), row_values(B, "Piutang usaha"), row_values(B, "Persediaan"),
                     row_values(B, "Jumlah aset lancar"))
ap, accr, adv = row_values(B, "Utang usaha"), row_values(B, "Akrual"), row_values(B, "Uang muka penjualan")
ocl, cl = row_values(B, "Liabilitas jangka pendek lainnya"), row_values(B, "Jumlah liabilitas jangka pendek")
ib1, lease1 = row_values(B, "Memo: Utang berbunga jatuh tempo"), row_values(B, "Liabilitas sewa jatuh tempo")
cfo, capex = row_values(C, "Arus kas bersih dari aktivitas operasi"), row_values(C, "Pembelian aset tetap")
div_nci = row_values(C, "kepentingan non-pengendali", startswith=False)
ebitda = block_row(P, "A. Comparable P&L", "EBITDA", "SIG")
rev = block_row(P, "A. Comparable P&L", "Net revenue", "SIG")
np_parent = row_values("IS (Build)", "Laba ke pemilik entitas induk")
i_cash, i_inv, i_ca = row_values(I, "Kas dan setara kas"), row_values(I, "Persediaan"), row_values(I, "Jumlah aset lancar")
i_ap, i_dep = row_values(I, "Utang usaha"), row_values(I, "Uang jaminan pelanggan")
i_cl, i_ib1 = row_values(I, "Jumlah liabilitas jangka pendek"), row_values(I, "Utang berbunga jatuh tempo")
i_cfo = row_values(I, "CFO setara SIG")
for name, vals in [("cash", cash), ("ar", ar), ("inv", inv), ("ca", ca), ("ap", ap), ("accr", accr), ("adv", adv),
                   ("ocl", ocl), ("cl", cl), ("ib1", ib1), ("lease1", lease1), ("cfo", cfo), ("capex", capex),
                   ("div_nci", div_nci), ("ebitda", ebitda), ("rev", rev), ("np_parent", np_parent),
                   ("i_cash", i_cash), ("i_inv", i_inv), ("i_ca", i_ca), ("i_ap", i_ap), ("i_dep", i_dep),
                   ("i_cl", i_cl), ("i_ib1", i_ib1), ("i_cfo", i_cfo)]:
    check_numbers(name, vals)

ratio = lambda a, b: [p / q for p, q in zip(a, b)]  # noqa: E731
cagr = lambda s: (s[4] / s[0]) ** 0.25 - 1  # noqa: E731
cov, quick, cfo_cl = ratio(cash, ib1), ratio([a - b for a, b in zip(ca, inv)], cl), ratio(cfo, cl)
i_cov, i_quick, i_cfo_cl = i_cash[4] / i_ib1[4], (i_ca[4] - i_inv[4]) / i_cl[4], i_cfo[4] / i_cl[4]
other = [a + b for a, b in zip(accr, ocl)]
i_other = i_cl[4] - i_ib1[4] - i_ap[4] - i_dep[4]
# current liabilities as % of current assets: SIG FY21, SIG FY25, INTP FY25
per = lambda part, base: part / base * 100  # noqa: E731
mix = {
    "Trade payables": [per(ap[0], ca[0]), per(ap[4], ca[4]), per(i_ap[4], i_ca[4])],
    "Short-term interest-bearing debt": [per(ib1[0], ca[0]), per(ib1[4], ca[4]), per(i_ib1[4], i_ca[4])],
    "Customer advances and deposits": [per(adv[0], ca[0]), per(adv[4], ca[4]), per(i_dep[4], i_ca[4])],
    "Other": [per(other[0], ca[0]), per(other[4], ca[4]), per(i_other, i_ca[4])],
}
tot = [sum(v[k] for v in mix.values()) for k in range(3)]
cash100 = [per(cash[0], ca[0]), per(cash[4], ca[4]), per(i_cash[4], i_ca[4])]
cr, i_cr = ca[4] / cl[4], i_ca[4] / i_cl[4]
cr_adj, i_cr_adj = ca[4] / (cl[4] - adv[4] - ib1[4]), i_ca[4] / (i_cl[4] - i_dep[4] - i_ib1[4])
cr_refi, cr_refi_ex_lease = ca[4] / (cl[4] - ib1[4]), ca[4] / (cl[4] - ib1[4] + lease1[4])
uses = ib1[4] - capex[4] + DIV_FY25 - div_nci[4]                       # capex and NCI dividends are negative
su, su_stress = (cash[4] + cfo[4]) / uses, (cash[4] + 0.7 * cfo[4]) / uses
div_keep = 0.5 * (np_parent[1] + np_parent[2]) / 2                     # 50% of FY22-23 average parent profit
days5 = rev[4] / 365 * 5
pay_share, pay_mult = mix["Trade payables"][1], mix["Trade payables"][1] / mix["Trade payables"][2]

# guards: every number on the slide and in the notes must hold in the data
assert all(abs(a + i + d + o - c) < 1 for a, i, d, o, c in zip(ap, ib1, adv, other, cl))
assert (x(cov[0]), x(cov[4]), x(i_cov)) == ("0.74x", "1.09x", "4.13x")
assert (x(quick[0]), x(quick[4]), x(i_quick)) == ("0.77x", "0.79x", "1.48x")
assert (x(cfo_cl[0]), x(cfo_cl[4]), x(i_cfo_cl)) == ("0.47x", "0.24x", "0.60x")
assert (pct(cagr(cov)), pct(cagr(quick)), pct(cagr(cfo_cl))) == ("+10.2%", "+0.5%", "\u201315.9%")
assert 0.22 < cfo_cl[4] < 0.27                                    # "covers only a quarter of short-term obligations"
assert abs(quick[4] - quick[0]) < 0.05                            # "quick ratio is back at its FY21 level"
assert cfo[4] / cfo[0] < 0.6 and ebitda[4] / ebitda[0] < 0.6      # "operating cash halved with EBITDA"
assert [round(v) for v in tot] == [90, 91, 54]                    # chart totals and headline Rp91 vs Rp54
assert [round(mix["Trade payables"][k]) for k in range(3)] == [49, 45, 14]
assert 40 < pay_share < 50 and pay_mult >= 3                      # "almost half ... three times INTP's share"
assert [round(v) for v in cash100] == [18, 27, 52]
assert round(i_cr_adj, 2) == 2.45 and round(cr_adj, 2) == 1.74   # notes: like-for-like adjusted ratios
assert round(cr_refi, 2) == 1.50 and round(cr_refi_ex_lease, 2) == 1.46
assert f"{su_stress:.1f}" == "1.3" and f"{su:.1f}" == "1.5"
assert round(div_keep / 1e6, 1) == 1.1 and round(days5 / 1e6, 1) == 0.5
assert round(PAYOUT_FY23 * 100) == 26 and (t(adv[0]), t(adv[4])) == ("0.1", "1.6")

# ---- slim deck with only slide 13 --------------------------------------------
prs = make_slim_deck(SRC, keep_slide_no=13)
slide = prs.slides[0]
clear_slide(slide)
ph = {s.placeholder_format.idx: s for s in slide.placeholders}
replace_runs_text(ph[0], "Liquidity")
replace_runs_text(ph[11], "Financial Analysis \u2013 Liquidity Analysis")

LX, LW = 40.6, 250.0
RX, RW = 305.0, 614.5

# ---- left: framing (wireframe metrics) + key takeaways -------------------------
add_column_header(slide, LX, 114, LW,
                  "Buffers held up, but operating cash now covers only a quarter of short-term obligations",
                  "Liquidity ratios, x")
add_kpi_table(slide, LX, 167, [74, 44, 44, 40, 48], ["SIG FY21", "SIG FY25", "CAGR", "INTP FY25"], [
    ("Cash \u00f7 debt due within a year", [x(cov[0]), x(cov[4]), pct(cagr(cov)), x(i_cov)]),
    ("Quick ratio", [x(quick[0]), x(quick[4]), pct(cagr(quick)), x(i_quick)]),
    ("Operating cash flow \u00f7 current liabilities", [x(cfo_cl[0]), (x(cfo_cl[4]), "\u25bc"), pct(cagr(cfo_cl)),
                                                      x(i_cfo_cl)]),
], row_h=32, header_h=14, label_size=9, value_size=9.5)
KY = 289.0
add_rect(slide, LX, KY, LW, 494 - KY, fill=PANEL_GREY, line=GREY, line_w=0.75, dash=True)
add_text(slide, LX + 10, KY + 8, LW - 20, 494 - KY - 14, [
    {"runs": [("Key takeaways", {"bold": True, "size": 12, "color": INK})], "space_after": 8},
    {"runs": [("1. Against what SIG holds: adequate. ", {"bold": True}),
              "Cash on hand still exceeds near-term debt, and the quick ratio is back at its FY21 level."],
     "space_after": 7},
    {"runs": [("2. Against what SIG generates: eroding. ", {"bold": True}),
              "Operating cash halved with EBITDA, so each year adds less to the buffer."], "space_after": 7},
    {"runs": [("3. Pressure point: supplier credit. ", {"bold": True}),
              "Suppliers fund almost half of current assets, three times INTP's share, so tighter terms would drain "
              "cash quickly."]},
], size=10.5)

# ---- right top: what holds the current ratio down (like for like) -------------
add_column_header(slide, RX, 114, RW,
                  f"SIG's current liabilities equal {tot[1]:.0f}% of its current assets (INTP {tot[2]:.0f}%), driven "
                  "mostly by supplier credit",
                  "Current liabilities as % of current assets", headline_h=16)
colors = [PAY_C, DEBT_C, ADV_C, OTH_C]
add_legend(slide, RX, 152, [(nm, col, "square") for nm, col in zip(mix, colors)], gap=10)
CW = 525.0
add_column_chart(
    slide, RX, 164, CW, 126, ["SIG FY21", "SIG FY25", "INTP FY25"],
    [{"name": nm, "values": vals, "color": col,
      "label_color": "FFFFFF" if col in (PAY_C, DEBT_C) else BODY, "label_size": 9,
      **({"label_points": [k for k, v in enumerate(vals) if v >= 3]})}
     for (nm, vals), col in zip(mix.items(), colors)],
    stacked=True, gap_width=120, y_min=0, y_max=104, number_format='0"%"',
    totals=[f"{v:.0f}%" for v in tot], total_pad=9, name="Chart liability mix")
BX, BWD = RX + RW - 74.5, 74.5
add_text(slide, BX - 6, 164, BWD + 12, 26, [[("Cash as % of current assets", {"bold": True})]], size=9,
         color=SKHA_BLUE, align="c", anchor="m")
add_summary_boxes(slide, BX, 190, BWD, "", [
    ("SIG FY21", f"{cash100[0]:.0f}%", "A9B9CC", INK, None),
    ("SIG FY25", f"{cash100[1]:.0f}%", SIG_NAVY, "FFFFFF", None),
    ("INTP FY25", f"{cash100[2]:.0f}%", INTP_BLUE, "FFFFFF", None),
], heading_h=1)

# ---- right bottom: recommendations ----------------------------------------------
add_text(slide, RX, 300, RW, 16, [[("Three moves SIG can start now to rebuild the buffer", {"bold": True})]], size=12,
         anchor="b")
add_line(slide, RX, 319, RX + RW, 319, SKHA_BLUE, 1.5)
lead = lambda head, body: [(head + " ", {"bold": True}), (body, {})]  # noqa: E731
add_text_table(slide, RX, 324, [("Recommendation", 330), ("Supporting evidence", RW - 330)], [
    [lead("Refinance near-term debt.", "Roll upcoming maturities into 3\u20135 year bonds or term loans, backed by a "
                                       "committed standby line"),
     "idAAA stable rating reaffirmed Aug 2026. Rp2.3 T of obligations repaid in Jan\u2013Jul 2026. Rp0.9 T "
     "refinanced at a 1 pp lower rate after the SMBR integration"],
    [lead("Tie dividends to cash.", "Propose to Danantara a payout policy linked to cash coverage, stepping down "
                                    "while cash is tight"),
     f"SIG paid out only {PAYOUT_FY23 * 100:.0f}% of FY2023 profit. Danantara sets dividends by each holding's "
     "needs (Feb 2026)"],
    [lead("Collect faster.", "Move more distributors to deposits or cash before delivery and shorten receivable "
                             "terms, without stretching suppliers"),
     f"Customer advances already rose from Rp{t(adv[0])} T to Rp{t(adv[4])} T since FY21"],
], header_h=15, row_h=50, size=9.5)

# ---- footer ----------------------------------------------------------------------
fill_text_frame(ph[10].text_frame, [
    "Source: SIG and INTP audited financial statements FY2021\u20132025, company disclosures"], size=8, color="7F7F7F")

# ---- speaker notes (Q&A material) ------------------------------------------------
notes = "\n".join([
    "Definitions: debt due within a year = interest-bearing debt (bank loans, bonds, short-term loans, leases) due "
    "within 12 months. Quick ratio = (current assets minus inventory) over current liabilities. Operating cash flow "
    "includes interest paid for both companies. INTP customer advances = customer deposits.",
    f"Like for like: removing customer advances and debt due within a year lifts SIG's current ratio from {x(cr)} to "
    f"{x(cr_adj)}, but INTP adjusted the same way is {x(i_cr_adj)} (reported {x(i_cr)}). The gap is mostly supplier "
    f"credit: trade payables equal {mix['Trade payables'][1]:.0f}% of current assets vs "
    f"{mix['Trade payables'][2]:.0f}% at INTP.",
    f"Operating cash Rp{t(cfo[0])} T to Rp{t(cfo[4])} T, EBITDA Rp{t(ebitda[0])} T to Rp{t(ebitda[4])} T.",
    f"Stress test: cash plus FY25 operating cash over debt due, capex and dividends = {su:.2f}x, {su_stress:.2f}x if "
    "operating cash falls 30%.",
    f"Impact estimates (not on the slide): refinancing all debt due within a year lifts the current ratio to about "
    f"{x(cr_refi)} ({x(cr_refi_ex_lease)} excluding leases). A 50% payout at the FY22 to FY23 average profit keeps "
    f"about Rp{t(div_keep)} T a year. Each 5 fewer receivable days frees about Rp{t(days5)} T.",
    "Evidence: PEFINDO affirmed idAAA stable on 11 Aug 2026. SIG cut obligations by Rp2.3 T in Jan to Jul 2026 "
    "(Stabilitas, Sep 2026). After the SMBR integration, Rp0.9 T was refinanced at a 1 pp lower rate (SIG corporate "
    "presentation, Mar 2024, p19). Payout 26.4% of FY2023 profit (Rp572 bn), 90.1% of FY2024, 100% of FY2025 "
    "(Rp190.8 bn, AGM 8 May 2026). Danantara: dividends no longer fixed, set by each holding's RKAP (Bisnis, Feb "
    "2026), and under Law 1/2025 they stay in Danantara rather than the state budget (RCTI+/Okezone, 30 Sep 2026).",
    "Considered and not recommended: group cash pooling (SIG has a BNI notional pool since 2019, useful but cash and "
    "debt fall together, so little liquidity effect). Scrip dividend (no recent IDX precedent, legal check needed). "
    "Asset recycling and receivable securitization (no SIG precedent, longer than 12 months, securitization is "
    "debt in economic terms). Shareholder loan from Danantara Asset Management (used only for a distressed SOE, "
    "Garuda 2025). Stretching suppliers through reverse factoring (Carillion 2018, called a hidden debt loophole by "
    "Fitch).",
])
slide.notes_slide.notes_text_frame.text = notes

# user preference: no semicolons anywhere on the slide (titles, headlines, KTW, notes)
slide_text = "\n".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
assert ";" not in slide_text and ";" not in notes, "semicolon found"

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print("saved", OUT, f"{OUT.stat().st_size / 1e6:.2f} MB")
print("mix", {k: [round(v, 1) for v in vals] for k, vals in mix.items()}, "tot", [round(v, 1) for v in tot],
      "cash100", [round(v, 1) for v in cash100])
print("adj cr", round(cr_adj, 2), round(i_cr_adj, 2), "refi", round(cr_refi, 2), round(cr_refi_ex_lease, 2),
      "s/u", round(su, 2), round(su_stress, 2), "div keep", round(div_keep / 1e6, 2), "5 days", round(days5 / 1e6, 2))
