"""Slide 10 v2: OPEX deep dive (core opex breakdown with CAGR, opex intensity vs INTP, KTW strip).

Delivers ONLY this slide as a slim .pptx (template from v1, one layout, Univers fonts only).
Every opex number comes from IS (Presentation) only (user rule): core opex = "SG&A excl. freight" of the
comparable P&L, selling excl. freight = that line minus reported G&A. Other operating income (expense) is net
and shown separately. The only outside figure is the benchmark: BPS sector PPI (same series as slide 9).
Run:  python build_slide10_opex.py
"""
import sys
from pathlib import Path

from pptx.enum.chart import XL_LABEL_POSITION as POS
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Pt

sys.path.insert(0, str(Path(__file__).parent))
from sig_data import PPI_KBLI23_Q4_YOY, YEARS, block_row, check_numbers, cumulative, row_values  # noqa: E402
from skha_helpers import (BODY, GREY, INK, INTP_BLUE, MUTED, PANEL_GREY, SIG_NAVY, add_column_chart,  # noqa: E402
                          add_column_header, add_legend, add_line, add_line_chart, add_rect, add_summary_boxes,
                          add_text, clear_slide, fill_text_frame, make_slim_deck, replace_runs_text, set_plot_area)

SRC = Path(r"D:\Download\Financial Analysis SIG (SMGR) v1.pptx")
OUT = Path(r"D:\Download\SIG_Financial_Analysis\slides\Slide_10_Opex_Deep_Dive_v2.pptx")

# Robustness check only (speaker notes): BPS headline CPI, Dec-on-Dec, same series as slide 9.
CPI_DEC_YOY = {2022: 0.0551, 2023: 0.0261, 2024: 0.0157, 2025: 0.0292}

SELL_C = "B4C3D3"     # selling excl. freight: light navy (SIG family), every segment labelled


def pct(v, d=1, sign=True):
    """Percent text with an en dash for negatives (no U+2212 dependency on Univers)."""
    s = f"{abs(v) * 100:.{d}f}%"
    return (("+" if v >= 0 else "\u2013") if sign else ("" if v >= 0 else "\u2013")) + s


def rp(v, d=2):
    """Signed Rp trillion text from Rp juta, e.g. '\u20130.29'."""
    return ("+" if v >= 0 else "\u2013") + f"{abs(v) / 1e6:.{d}f}"


# ---- data: IS (Presentation) only ------------------------------------------------
P, A = "IS (Presentation)", "A. Comparable P&L"
rev, i_rev = block_row(P, A, "Net revenue", "SIG"), block_row(P, A, "Net revenue", "INTP")
opex = [-v for v in block_row(P, A, "SG&A excl. freight", "SIG")]
i_opex = [-v for v in block_row(P, A, "SG&A excl. freight", "INTP")]
oth, i_oth = (block_row(P, A, "Other operating income (expense)", "SIG"),
              block_row(P, A, "Other operating income (expense)", "INTP"))
gp = block_row(P, A, "Gross profit after freight", "SIG")
ebit = block_row(P, A, "Operating profit (EBIT)", "SIG")
deliv = [-v for v in block_row(P, A, "Total delivered cost", "SIG")]
vol, i_vol = block_row(P, "B. Per ton", "Volume", "SIG"), block_row(P, "B. Per ton", "Volume", "INTP")
sell_rep = [-v for v in row_values(P, "Beban penjualan")]           # reported, incl. freight
ga = [-v for v in row_values(P, "Beban umum & administrasi")]
cogs = [-v for v in row_values(P, "Total COGS")]
oth_rep = row_values(P, "Pendapatan (beban) lainnya")
for name, vals in [("rev", rev), ("i_rev", i_rev), ("opex", opex), ("i_opex", i_opex), ("oth", oth),
                   ("i_oth", i_oth), ("gp", gp), ("ebit", ebit), ("deliv", deliv), ("vol", vol), ("i_vol", i_vol),
                   ("sell_rep", sell_rep), ("ga", ga), ("cogs", cogs), ("oth_rep", oth_rep)]:
    check_numbers(name, vals)

sell = [o - g for o, g in zip(opex, ga)]                    # selling excl. freight
freight = [s - x for s, x in zip(sell_rep, sell)]           # freight booked in reported selling
cagr = lambda s: (s[4] / s[0]) ** 0.25 - 1  # noqa: E731
ppi = cumulative(PPI_KBLI23_Q4_YOY)
cagr_s, cagr_i, cagr_ppi, cagr_cpi = cagr(opex), cagr(i_opex), ppi ** 0.25 - 1, cumulative(CPI_DEC_YOY) ** 0.25 - 1
cagr_sell, cagr_ga = cagr(sell), cagr(ga)
incl, i_incl = [o - x for o, x in zip(opex, oth)], [o - x for o, x in zip(i_opex, i_oth)]   # opex incl. other op
cagr_incl, cagr_i_incl = cagr(incl), cagr(i_incl)
ratio, i_ratio = [o / r for o, r in zip(opex, rev)], [o / r for o, r in zip(i_opex, i_rev)]
gap = [a - b for a, b in zip(ratio, i_ratio)]
avg_s, avg_i = sum(ratio) / 5, sum(i_ratio) / 5
m = lambda s, k: s[k] / rev[k]  # noqa: E731
# EBIT margin bridge FY21 -> FY25, each term = impact on EBIT margin (negative = drag)
d_ebit, d_gp, d_opex, d_oth = (m(ebit, 4) - m(ebit, 0), m(gp, 4) - m(gp, 0), -(ratio[4] - ratio[0]),
                               m(oth, 4) - m(oth, 0))
per_t, i_per_t = opex[4] / vol[4], i_opex[4] / i_vol[4]
sell_rep_chg = sell_rep[4] / sell_rep[0] - 1
freight_share = (freight[0] - freight[4]) / (sell_rep[0] - sell_rep[4])
rep_ratio = [(s + g) / r for s, g, r in zip(sell_rep, ga, rev)]

# guards: every claim on the slide must hold in the data
assert oth == oth_rep                                                          # comparable = reported other op
assert all(abs(f - (d - c)) < 1 for f, d, c in zip(freight, deliv, cogs))      # freight in selling = deliv - COGS
assert abs(d_gp + d_opex + d_oth - d_ebit) < 1e-9                              # margin bridge closes
assert round(cagr_s * 100, 1) == 1.5 and round(cagr_i * 100, 1) == 7.2        # "+1.5% a year", "INTP +7.2%"
assert round(cagr_sell * 100, 1) == -1.2 and round(cagr_ga * 100, 1) == 2.2   # legend CAGRs
assert round(cagr_ppi * 100, 1) == 2.9 and cagr_s < cagr_ppi < cagr_i          # "below sector producer prices"
assert abs(cagr_incl - cagr_ppi) < 0.001 and round(cagr_incl * 100, 1) == 2.9  # "in line with ... (2.9%)"
assert round(cagr_i_incl * 100, 1) == 12.5                                     # "far below INTP (12.5%)"
assert round(gap[0] * 100, 1) == 3.7 and round(gap[4] * 100, 1) == 4.1         # "3.7 -> 4.1 pp", "~4 pp"
assert round(-d_opex * 100, 1) == 1.2 and round(-d_ebit * 100, 1) == 10.6     # "1.2 of the 10.6 pp"
assert round(-oth[4] / ebit[4], 2) == 0.23                                     # other op FY25 = 23% of EBIT
assert round(oth[4] / 1e6, 2) == -0.29                                         # "net expense Rp0.29 T"

# ---- slim deck with only slide 10 --------------------------------------------
prs = make_slim_deck(SRC, keep_slide_no=10)
slide = prs.slides[0]
clear_slide(slide)
ph = {s.placeholder_format.idx: s for s in slide.placeholders}
replace_runs_text(ph[0], "OPEX deep dive")
replace_runs_text(ph[11], "Financial Analysis \u2013 Profitability Analysis")

LX, LW = 40.6, 429.4
RX, RW = 490.0, 429.5

# ---- left: core opex breakdown with CAGR ---------------------------------------
add_column_header(slide, LX, 114, LW,
                  f"SIG core opex barely moved (Rp{opex[0] / 1e6:.1f} T \u2192 Rp{opex[4] / 1e6:.1f} T, "
                  f"{pct(cagr_s)} a year), well below INTP's {pct(cagr_i)}",
                  "Core opex by type, Rp T (selling excl. freight + G&A)")
add_legend(slide, LX, 166, [(f"G&A (CAGR {pct(cagr_ga)})", SIG_NAVY, "square"),
                            (f"Selling excl. freight (CAGR {pct(cagr_sell)})", SELL_C, "square")], gap=12)
CY, CH = 178.0, 192.0
PX0, PW, PY0, PH = 0.03, 0.94, 0.20, 0.66            # inner plot area as fractions of the chart frame
gf = add_column_chart(
    slide, LX, CY, LW, CH, YEARS,
    [{"name": "G&A", "values": [v / 1e6 for v in ga], "color": SIG_NAVY, "label_color": "FFFFFF", "label_size": 9},
     {"name": "Selling excl. freight", "values": [v / 1e6 for v in sell], "color": SELL_C, "label_size": 8.5}],
    stacked=True, gap_width=55, y_min=0, y_max=5.0, number_format="0.00",
    totals=[f"{v / 1e6:.1f}" for v in opex], total_pad=0.45, name="Chart SIG core opex")
set_plot_area(gf.chart, PX0, PY0, PW, PH)
xc = [LX + LW * (PX0 + PW * (i + 0.5) / 5) for i in range(5)]   # category centres on the slide
# CAGR arrow across the bars
ay = CY + CH * PY0 - 13
add_line(slide, xc[0], ay, xc[4], ay, SIG_NAVY, 1.0, tail="triangle")
cagr_box = add_rect(slide, (xc[0] + xc[4]) / 2 - 58, ay - 8, 116, 16, fill=SIG_NAVY)
fill_text_frame(cagr_box.text_frame, [[(f"CAGR FY21\u201325  {pct(cagr_s)}", {"bold": True})]], size=9,
                color="FFFFFF", align="c")
cagr_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
# supporting metric: other operating income (expense), net, aligned with the bars
SY = CY + CH + 2
add_rect(slide, LX, SY, LW, 28, fill=PANEL_GREY)
add_text(slide, LX + 6, SY + 2, LW - 12, 11,
         [[("Other operating income (expense), net, Rp T", {"bold": True}),
           (" \u2013 shown separately, not in core opex", {})]], size=8, color=MUTED)
for x, v in zip(xc, oth):
    add_text(slide, x - 24, SY + 14, 48, 12, [[(rp(v), {"bold": True})]], size=9, color=BODY, align="c")

# ---- right: opex intensity benchmark (reference-chart style) -------------------
add_column_header(slide, RX, 114, RW,
                  f"SIG spends ~{gap[4] * 100:.0f} pp more of revenue on opex than INTP, a structural gap that "
                  f"barely moved since FY21 ({gap[0] * 100:.1f} \u2192 {gap[4] * 100:.1f} pp)",
                  "Core opex as % of revenue, same definition for both")
CHW = 300.0
add_line_chart(slide, RX, 166, CHW, 232, YEARS,
               [{"name": "INTP", "values": i_ratio, "color": INTP_BLUE, "width": 1.5,
                 "label_all": POS.BELOW, "label_fill": True},
                {"name": "SIG", "values": ratio, "color": SIG_NAVY, "width": 1.5,
                 "label_all": POS.ABOVE, "label_fill": True}],
               number_format="0.0%", y_min=0.04, y_max=0.15, name="Chart opex intensity")
add_legend(slide, RX + CHW + 4, 240, [("SIG", SIG_NAVY, "line"), ("INTP", INTP_BLUE, "line")],
           line_w=1.5, vertical=True)
add_summary_boxes(slide, RX + RW - 62, 172, 62, "Avg 21\u201325", [
    ("SIG", pct(avg_s, sign=False), SIG_NAVY, "FFFFFF", None),
    ("INTP", pct(avg_i, sign=False), INTP_BLUE, "FFFFFF", None),
])
add_summary_boxes(slide, RX + RW - 62, 262, 62, "Gap, pp", [
    ("FY21", f"{gap[0] * 100:.1f}", PANEL_GREY, BODY, (MUTED, True)),
    ("FY25", f"{gap[4] * 100:.1f}", PANEL_GREY, BODY, (MUTED, True)),
])

# ---- KTW strip under both charts ---------------------------------------------
KY, KH = 408.0, 86.0
add_rect(slide, LX, KY, RX + RW - LX, KH, fill=PANEL_GREY, line=GREY, line_w=0.75, dash=True)
add_text(slide, LX + 10, KY + 5, 200, 14, [[("Key takeaways", {"bold": True})]], size=11, color=INK)
cols = [
    [("Not the problem, stays at BAU. ", {"bold": True}),
     f"Core opex grew {pct(cagr_s, sign=False)} a year, below sector producer prices ({pct(cagr_ppi, sign=False)}) "
     f"and INTP ({pct(cagr_i, sign=False)}), and explains only {-d_opex * 100:.1f} of the {-d_ebit * 100:.1f} pp "
     "fall in SIG's EBIT margin."],
    [("The gap is structural. ", {"bold": True}),
     f"SIG runs a holding with many entities and multiple brands. The gap barely moved ({gap[0] * 100:.1f} "
     f"\u2192 {gap[4] * 100:.1f} pp), so it did not cause the decline, but streamlining can narrow it."],
    [("Devil's advocate check. ", {"bold": True}),
     f"Adding other operating items (FY25 net expense Rp{-oth[4] / 1e6:.2f} T) lifts SIG's opex growth to "
     f"{pct(cagr_incl, sign=False)} a year, in line with sector producer prices ({pct(cagr_ppi, sign=False)}) "
     f"and far below INTP ({pct(cagr_i_incl, sign=False)})."],
]
cw, cgap = (RX + RW - LX - 20 - 2 * 16) / 3, 16
for i, runs in enumerate(cols):
    add_text(slide, LX + 10 + i * (cw + cgap), KY + 23, cw, KH - 27, [runs], size=10.5)

# ---- footer --------------------------------------------------------------------
source = ph[10]
_l, _t, _h = source.left, source.top, source.height   # copy layout position before resizing (see slide 9)
source.left, source.top, source.width, source.height = _l, _t, Pt(790) - _l, _h
fill_text_frame(source.text_frame, [
    "Note: core opex = selling expenses excl. freight + G&A, comparable basis from IS (Presentation), with freight in "
    "delivered cost for both. Other operating income (expense) is net and shown separately. Sector PPI = BPS producer "
    "price index, non-metallic mineral products (KBLI 23), Q4 to Q4. Source: SIG and INTP audited financial "
    "statements FY2021\u20132025, BPS, Skha analysis"], size=8, color="7F7F7F")

# ---- speaker notes -------------------------------------------------------------
notes = "\n".join([
    "All opex figures are from IS (Presentation). Core opex = SG&A excl. freight from the comparable P&L. Selling "
    "excl. freight = SG&A excl. freight minus reported G&A.",
    f"Why not reported figures: SIG's reported selling expenses fell {pct(sell_rep_chg, 0)} (Rp{sell_rep[0] / 1e6:.2f} T "
    f"\u2192 Rp{sell_rep[4] / 1e6:.2f} T), but {freight_share * 100:.0f}% of the drop is freight moving out of "
    f"selling (Rp{freight[0] / 1e6:.2f} T \u2192 Rp{freight[4] / 1e6:.2f} T). On a reported basis SIG's SG&A ratio "
    f"falls from {pct(rep_ratio[0], sign=False)} to {pct(rep_ratio[4], sign=False)}, a false saving. INTP books all "
    "freight in selling, so only the comparable basis is like for like.",
    f"Devil's advocate on other operating items: the line is net (income and expense mixed) and flips sign "
    f"(SIG {', '.join(rp(v) for v in oth)} Rp T). Including it, SIG opex grows {pct(cagr_incl)} a year vs INTP "
    f"{pct(cagr_i_incl)}. The FY25 net expense equals {-oth[4] / ebit[4] * 100:.0f}% of EBIT, so it is shown, but "
    "separately.",
    f"Benchmark: sector PPI (BPS, non-metallic mineral products) {pct(ppi - 1)} cumulative, {pct(cagr_ppi)} a year. "
    f"It asks whether overhead grew faster than what the sector can charge. CPI ({pct(cagr_cpi)} a year) gives the "
    "same answer.",
    f"Per ton FY25: SIG Rp{per_t:.0f}k vs INTP Rp{i_per_t:.0f}k.",
    f"EBIT margin {pct(m(ebit, 0), sign=False)} \u2192 {pct(m(ebit, 4), sign=False)}: gross margin after freight "
    f"{pct(d_gp).replace('%', ' pp')}, opex {pct(d_opex).replace('%', ' pp')}, other operating "
    f"{pct(d_oth).replace('%', ' pp')}.",
    "Structural gap (interpretation): SIG is a holding with many entities, now being streamlined (several entities "
    "liquidated May to July 2026 per company disclosures), and runs multiple cement brands (SIG corporate "
    "presentation). The gap has not widened, so it explains the level, not the decline.",
])
slide.notes_slide.notes_text_frame.text = notes

# user preference: no semicolons anywhere on the slide (titles, headlines, KTW, notes)
slide_text = "\n".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
assert ";" not in slide_text and ";" not in notes, "semicolon found"

OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print("saved", OUT, f"{OUT.stat().st_size / 1e6:.2f} MB")
print("opex", [round(v / 1e6, 2) for v in opex], "sell", [round(v / 1e6, 2) for v in sell], "ga",
      [round(v / 1e6, 2) for v in ga], "oth", [round(v / 1e6, 2) for v in oth])
print("cagr s/i/ppi/cpi", round(cagr_s, 4), round(cagr_i, 4), round(cagr_ppi, 4), round(cagr_cpi, 4), "incl",
      round(cagr_incl, 4), round(cagr_i_incl, 4), "gap", [round(g, 4) for g in gap], "avg", round(avg_s, 4),
      round(avg_i, 4))
print("xc", [round(x, 1) for x in xc], "arrow y", round(ay, 1))
