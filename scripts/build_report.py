"""Build a GPS Impact-branded digital competitive report from an Ad Hawk data pull.

Usage:
    python3 build_report.py --data-dir DIR --title TITLE --output OUT.xlsx [--today YYYY-MM-DD] [--dvr-tab]

Expects DIR to contain two files produced by parsing the Ad Hawk Google Sheet
(see the "Raw Creative" and daily-spend tabs):
    raw_creative_clean.json  - one record per creative, with first_ran/last_ran
    daily_spend_clean.json   - one record per advertiser per day

--dvr-tab adds a Democrat vs. Republican summary tab; only meaningful for a
two-party general election (skip it for primaries or nonpartisan races).

The tab-builder functions below (add_competitive_report_tab,
add_creative_timeline_tab, add_dvr_tab) are reused as-is by
build_multi_district_report.py to assemble one workbook covering many
races (one race's worth of rows in, one styled tab out, every time).
"""
import argparse
import json
import datetime
import os
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.formatting.rule import ColorScaleRule

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_WHITE = os.path.join(SCRIPT_DIR, "..", "assets", "logos", "GPSImpact_White_Horizontal_2026.png")

BANNER_ROW1_PT = 34
BANNER_ROW2_PT = 18
BANNER_HEIGHT_PX = (BANNER_ROW1_PT + BANNER_ROW2_PT) * 4 / 3  # pt -> px

# =========================================================================
# GPS IMPACT BRAND — colors pulled from GPS Impact's own AdImpact-based
# template (Competitive Digital Report tab): each party gets a 3-step tint
# ramp (candidate block -> subtotal -> candidate total) plus a solid brand
# color for the party-level total row. Ramp ratios (blend-to-white amounts
# 0.90 / 0.70 / 0.615) were reverse-engineered from that file's own R/D
# colors. Parties without a defined brand color fall back to a neutral
# gray ramp using the same ratios.
# =========================================================================
NAVY      = '323b51'   # GPS primary - Navy
BLUE      = '3d6a91'   # GPS primary - Blue (Democrat)
RED       = 'de5e4e'   # GPS secondary - Red (Republican)
ORANGE    = 'f7a747'   # GPS extended - Non-Partisan orange
ORANGE_DK = 'd37518'   # GPS extended - Non-Partisan orange (dark/solid)
GREY_TXT  = '666666'
FOOTER_GREY = '6B7280'

PARTY_BASE  = {'D': BLUE, 'R': RED, 'NP': ORANGE}
PARTY_SOLID = {'D': BLUE, 'R': RED, 'NP': ORANGE_DK}
PARTY_LABEL = {'D': 'DEMOCRAT PARTY TOTAL', 'R': 'REPUBLICAN PARTY TOTAL', 'NP': 'NON-PARTISAN PARTY TOTAL',
               'I': 'INDEPENDENT PARTY TOTAL'}
DEFAULT_BASE, DEFAULT_SOLID = '999999', '595959'

TITLE_FONT = 'Superior Title'
BODY_FONT  = 'Figtree'

MONEY = '$#,##0;-$#,##0;""'
MONEY_TOT = '$#,##0'

hdr_font    = Font(name=BODY_FONT, size=9, bold=True, color='FFFFFF')
hdr_fill    = PatternFill('solid', fgColor=NAVY)
title_font  = Font(name=TITLE_FONT, size=14, bold=True, color='FFFFFF')
note_font   = Font(name=BODY_FONT, size=8, italic=True, color=GREY_TXT)
footer_font = Font(name=BODY_FONT, size=8, italic=True, color=FOOTER_GREY)
cell_font   = Font(name=BODY_FONT, size=9, color=NAVY)
adv_font    = Font(name=BODY_FONT, size=9, bold=True, color=NAVY)
tot_font    = Font(name=BODY_FONT, size=9, bold=True, color=NAVY)
party_tot_font = Font(name=BODY_FONT, size=10, bold=True, color='FFFFFF')
grand_tot_font = Font(name=BODY_FONT, size=10, bold=True, color='FFFFFF')

BORDER_TOTAL = Border(top=Side(style='medium', color=NAVY), bottom=Side(style='medium', color=NAVY))
BORDER_GRAND = Border(top=Side(style='thick', color=NAVY), bottom=Side(style='double', color=NAVY))
BORDER_LIVE = Border(top=Side(style='thin', color=NAVY), bottom=Side(style='thin', color=NAVY),
                      left=Side(style='thin', color=NAVY), right=Side(style='thin', color=NAVY))


def tint(hex_color, amount):
    r = int(hex_color[0:2], 16); g = int(hex_color[2:4], 16); b = int(hex_color[4:6], 16)
    r = round(r + (255 - r) * amount)
    g = round(g + (255 - g) * amount)
    b = round(b + (255 - b) * amount)
    return f'{r:02x}{g:02x}{b:02x}'


def party_colors(party):
    base = PARTY_BASE.get(party, DEFAULT_BASE)
    solid = PARTY_SOLID.get(party, DEFAULT_SOLID)
    return tint(base, 0.90), tint(base, 0.70), tint(base, 0.615), solid  # leaf, subtotal, candidate-total, solid


def week_start_of(d):
    """Monday-of-week shifted to a Tuesday-start convention."""
    return d - datetime.timedelta(days=(d.weekday() - 1) % 7)


def weeks_covering(dates, today):
    """Every Tuesday-start week from the earliest date given through today."""
    min_week_start = week_start_of(min(dates))
    this_week_start = week_start_of(today)
    weeks = []
    w = min_week_start
    while w <= this_week_start:
        weeks.append(w)
        w += datetime.timedelta(days=7)
    return weeks


def style_row(ws, row, ncols, font=None, fill=None, border=None, start=1):
    for c in range(start, ncols + 1):
        cell = ws.cell(row=row, column=c)
        if font: cell.font = font
        if fill: cell.fill = fill
        if border: cell.border = border


def add_logo(ws, height_px=26):
    """Place the logo in the top-left, vertically centered across the
    navy title banner (rows 1-2)."""
    img = XLImage(LOGO_WHITE)
    aspect = img.width / img.height
    img.height = height_px
    img.width = height_px * aspect
    row_off_px = max(0, (BANNER_HEIGHT_PX - height_px) / 2)
    marker = AnchorMarker(col=0, colOff=0, row=0, rowOff=pixels_to_EMU(row_off_px))
    size = XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height))
    img.anchor = OneCellAnchor(_from=marker, ext=size)
    ws.add_image(img)


def add_footer(ws, row, ncols):
    ws.cell(row=row, column=1, value="Report prepared by GPS Impact  |  Confidential").font = footer_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)


def _advertiser_order(rows2):
    """Party blocks ordered by total party spend desc; within each party,
    advertisers ordered by spend desc. Shared by every tab so the same
    race always lists advertisers in the same order."""
    adv_party = {}
    adv_total = defaultdict(float)
    for r in rows2:
        adv_party[r['advertiser']] = r['party']
        adv_total[r['advertiser']] += r['spend']

    party_total_spend = defaultdict(float)
    for adv, tot in adv_total.items():
        party_total_spend[adv_party[adv]] += tot
    parties_ordered = sorted(party_total_spend.keys(), key=lambda p: -party_total_spend[p])
    advertisers_by_spend = sorted(adv_total.keys(), key=lambda a: -adv_total[a])
    advertisers_ordered = []
    for party in parties_ordered:
        advertisers_ordered.extend([a for a in advertisers_by_spend if adv_party[a] == party])
    return advertisers_ordered, adv_party, adv_total, party_total_spend


def add_competitive_report_tab(wb, sheet_name, banner_title, weeks, rows2):
    """One race's spend rows -> one styled "Competitive Digital Report"-
    shaped tab (candidate/committee -> party -> platform -> weekly spend),
    added to workbook wb. Returns per-race stats a Summary tab needs:
    {platform: {'total':, 'GOP':, 'DEM':}}, plus the race's grand total."""
    advertisers_ordered, adv_party, adv_total, party_total_spend = _advertiser_order(rows2)

    weekly = defaultdict(lambda: defaultdict(float))
    platforms_by_adv = defaultdict(set)
    for r in rows2:
        wk = week_start_of(r['target_date'])
        weekly[(r['advertiser'], r['source_platform'])][wk] += r['spend']
        platforms_by_adv[r['advertiser']].add(r['source_platform'])

    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    HEADERS = ["CANDIDATE / COMMITTEE", "PARTY", "PLATFORM", "TOTAL SPEND"] + [wk.strftime("%m/%d/%Y") for wk in weeks]
    ncols = len(HEADERS)
    WK_COL0 = 5

    ws.row_dimensions[1].height = BANNER_ROW1_PT
    ws.row_dimensions[2].height = BANNER_ROW2_PT
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=4)
    tcell = ws.cell(row=1, column=1, value=f"   {banner_title}")
    tcell.font = title_font
    tcell.alignment = Alignment(horizontal='right', vertical='center')
    ws.merge_cells(start_row=1, start_column=WK_COL0, end_row=1, end_column=ncols)
    style_row(ws, 1, ncols, fill=hdr_fill)
    style_row(ws, 2, ncols, fill=hdr_fill)
    add_logo(ws)

    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    r = 4

    def flush_party_total(party, r):
        light, midt, candt, solid = party_colors(party)
        solid_fill = PatternFill('solid', fgColor=solid)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.cell(row=r, column=1, value=PARTY_LABEL.get(party, f"{party} TOTAL"))
        tot = sum(sum(weekly[(a, p)].values()) for a in advertisers_ordered if adv_party[a] == party for p in platforms_by_adv[a])
        ws.cell(row=r, column=4, value=tot).number_format = MONEY_TOT
        for ci, wk in enumerate(weeks, start=WK_COL0):
            v = sum(weekly[(a, p)].get(wk, 0) for a in advertisers_ordered if adv_party[a] == party for p in platforms_by_adv[a])
            ws.cell(row=r, column=ci, value=v).number_format = MONEY_TOT
        style_row(ws, r, ncols, font=party_tot_font, fill=solid_fill, border=BORDER_TOTAL)
        return r + 2

    prev_party = None
    for adv in advertisers_ordered:
        party = adv_party.get(adv, '')
        if prev_party is not None and party != prev_party:
            r = flush_party_total(prev_party, r)
        prev_party = party

        light, midt, candt, solid = party_colors(party)
        light_fill = PatternFill('solid', fgColor=light)
        cand_fill  = PatternFill('solid', fgColor=candt)

        block_start = r
        plats = sorted(platforms_by_adv[adv], key=lambda p: -sum(weekly[(adv, p)].values()))

        for p in plats:
            wk_spend = weekly[(adv, p)]
            total = sum(wk_spend.values())
            style_row(ws, r, ncols, font=cell_font, fill=light_fill)
            ws.cell(row=r, column=3, value=p)
            ws.cell(row=r, column=4, value=total).number_format = MONEY
            for ci, wk in enumerate(weeks, start=WK_COL0):
                v = wk_spend.get(wk, 0)
                ws.cell(row=r, column=ci, value=v).number_format = MONEY
            r += 1

        ws.cell(row=r, column=3, value=f"{adv} Total")
        ws.cell(row=r, column=4, value=adv_total[adv]).number_format = MONEY_TOT
        for ci, wk in enumerate(weeks, start=WK_COL0):
            v = sum(weekly[(adv, p)].get(wk, 0) for p in plats)
            ws.cell(row=r, column=ci, value=v).number_format = MONEY_TOT
        style_row(ws, r, ncols, font=tot_font, fill=cand_fill, border=BORDER_TOTAL)
        r += 1

        ws.merge_cells(start_row=block_start, start_column=1, end_row=r - 1, end_column=1)
        ws.merge_cells(start_row=block_start, start_column=2, end_row=r - 1, end_column=2)
        ws.cell(row=block_start, column=1, value=adv).font = adv_font
        ws.cell(row=block_start, column=2, value=party).font = adv_font
        ws.cell(row=block_start, column=1).fill = light_fill
        ws.cell(row=block_start, column=2).fill = light_fill
        for cc in (1, 2):
            ws.cell(row=block_start, column=cc).alignment = Alignment(vertical='top')

        r += 1

    if prev_party is not None:
        r = flush_party_total(prev_party, r)

    grand_fill = PatternFill('solid', fgColor=NAVY)
    ws.cell(row=r, column=1, value="GRAND TOTAL")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=4, value=sum(adv_total.values())).number_format = MONEY_TOT
    for ci, wk in enumerate(weeks, start=WK_COL0):
        v = sum(weekly[(a, p)].get(wk, 0) for a in advertisers_ordered for p in platforms_by_adv[a])
        ws.cell(row=r, column=ci, value=v).number_format = MONEY_TOT
    style_row(ws, r, ncols, font=grand_tot_font, fill=grand_fill, border=BORDER_GRAND)
    r += 2
    add_footer(ws, r, ncols)

    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 38
    ws.column_dimensions['D'].width = 13
    for c in range(WK_COL0, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 12
    ws.freeze_panes = get_column_letter(WK_COL0) + "4"
    ws.row_dimensions[3].height = 30

    # per-platform GOP/DEM/total, for a Summary tab rolling up many races
    platform_stats = defaultdict(lambda: {'total': 0.0, 'GOP': 0.0, 'DEM': 0.0})
    for r_ in rows2:
        stats = platform_stats[r_['source_platform']]
        stats['total'] += r_['spend']
        if r_['party'] == 'R':
            stats['GOP'] += r_['spend']
        elif r_['party'] == 'D':
            stats['DEM'] += r_['spend']
    return platform_stats


def add_creative_timeline_tab(wb, sheet_name, weeks, rows1, rows2_for_order=None):
    """One race's creative rows -> one styled "Creative Timeline" tab
    (weekly Gantt-style flight shading), added to workbook wb."""
    advertisers_ordered, adv_party, _, _ = _advertiser_order(rows2_for_order if rows2_for_order is not None else rows1)

    ws2 = wb.create_sheet(sheet_name)
    ws2.sheet_view.showGridLines = False
    CT_HEADERS = ["CANDIDATE / COMMITTEE", "CREATIVE", "PLATFORM", "PARTY", "PURPOSE", "TONE",
                  "TOPIC", "TOTAL SPEND", "FIRST RAN", "LAST RAN", "AD LIBRARY LINK"]
    ct_week_labels = [wk.strftime("%-m/%-d") for wk in weeks]
    ncols2 = len(CT_HEADERS) + len(weeks)

    ws2.row_dimensions[1].height = BANNER_ROW1_PT
    ws2.row_dimensions[2].height = BANNER_ROW2_PT
    ws2.merge_cells(start_row=1, start_column=1, end_row=2, end_column=len(CT_HEADERS))
    tcell2 = ws2.cell(row=1, column=1, value="   CREATIVE TIMELINE")
    tcell2.font = title_font
    tcell2.alignment = Alignment(horizontal='right', vertical='center')
    ws2.merge_cells(start_row=1, start_column=len(CT_HEADERS) + 1, end_row=1, end_column=ncols2)
    style_row(ws2, 1, ncols2, fill=hdr_fill)
    style_row(ws2, 2, ncols2, fill=hdr_fill)
    add_logo(ws2)

    for c, h in enumerate(CT_HEADERS, start=1):
        cell = ws2.cell(row=3, column=c, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for ci, lbl in enumerate(ct_week_labels, start=len(CT_HEADERS) + 1):
        cell = ws2.cell(row=3, column=ci, value=lbl)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

    creatives_by_adv = defaultdict(list)
    for cre in rows1:
        creatives_by_adv[cre['advertiser']].append(cre)

    # advertisers that only appear in creative rows, not in the spend-based
    # ordering (e.g. a committee with creatives logged but $0 spend so far)
    extra_advs = [a for a in creatives_by_adv if a not in adv_party]
    order = advertisers_ordered + sorted(extra_advs)

    r = 4
    for adv in order:
        creas = sorted(creatives_by_adv.get(adv, []), key=lambda c: c['first_ran'])
        if not creas:
            continue
        party = adv_party.get(adv) or creas[0].get('party', '')
        light, midt, candt, solid = party_colors(party)
        light_fill = PatternFill('solid', fgColor=light)
        mid_fill = PatternFill('solid', fgColor=midt)

        block_start = r
        for cre in creas:
            snippet = cre['transcript'].strip().replace('\n', ' ')
            snippet = (snippet[:90] + '…') if len(snippet) > 90 else snippet
            ws2.cell(row=r, column=2, value=snippet).font = cell_font
            ws2.cell(row=r, column=3, value=cre['source_platform']).font = cell_font
            ws2.cell(row=r, column=4, value=cre['party']).font = cell_font
            ws2.cell(row=r, column=5, value=cre['purpose']).font = cell_font
            ws2.cell(row=r, column=6, value=cre['sentiment']).font = cell_font
            ws2.cell(row=r, column=7, value=cre['topic']).font = cell_font
            sc = ws2.cell(row=r, column=8, value=cre['spend'])
            sc.number_format = MONEY_TOT
            sc.font = cell_font
            ws2.cell(row=r, column=9, value=cre['first_ran'].strftime('%m/%d/%Y')).font = cell_font
            ws2.cell(row=r, column=10, value=cre['last_ran'].strftime('%m/%d/%Y')).font = cell_font
            lc = ws2.cell(row=r, column=11, value="View Ad")
            lc.font = Font(name=BODY_FONT, size=9, color='0563C1', underline='single')
            lc.hyperlink = cre['link']
            lc.alignment = Alignment(horizontal='center')

            for ci, wk in enumerate(weeks, start=len(CT_HEADERS) + 1):
                week_end = wk + datetime.timedelta(days=6)
                if cre['first_ran'] <= week_end and cre['last_ran'] >= wk:
                    live_cell = ws2.cell(row=r, column=ci)
                    live_cell.fill = mid_fill
                    live_cell.border = BORDER_LIVE
            r += 1

        ws2.merge_cells(start_row=block_start, start_column=1, end_row=r - 1, end_column=1)
        ws2.cell(row=block_start, column=1, value=adv).font = adv_font
        ws2.cell(row=block_start, column=1).alignment = Alignment(vertical='top')
        for rr in range(block_start, r):
            ws2.cell(row=rr, column=1).fill = light_fill
        r += 1

    add_footer(ws2, r + 1, ncols2)

    ws2.column_dimensions['A'].width = 24
    ws2.column_dimensions['B'].width = 45
    ws2.column_dimensions['C'].width = 10
    ws2.column_dimensions['D'].width = 8
    ws2.column_dimensions['E'].width = 14
    ws2.column_dimensions['F'].width = 10
    ws2.column_dimensions['G'].width = 16
    ws2.column_dimensions['H'].width = 12
    ws2.column_dimensions['I'].width = 12
    ws2.column_dimensions['J'].width = 12
    ws2.column_dimensions['K'].width = 14
    for c in range(len(CT_HEADERS) + 1, ncols2 + 1):
        ws2.column_dimensions[get_column_letter(c)].width = 6
    ws2.freeze_panes = get_column_letter(len(CT_HEADERS) + 1) + "4"
    ws2.row_dimensions[3].height = 30


def add_dvr_tab(wb, weeks, rows2, party_total_spend):
    ws3 = wb.create_sheet("D vs R")
    ws3.sheet_view.showGridLines = False
    ws3.row_dimensions[1].height = BANNER_ROW1_PT
    ws3.row_dimensions[2].height = BANNER_ROW2_PT
    ncols3 = 4  # WEEK, DEMOCRAT, REPUBLICAN, DELTA

    ws3.merge_cells(start_row=1, start_column=1, end_row=2, end_column=ncols3)
    tcell3 = ws3.cell(row=1, column=1, value="  D VS R")
    tcell3.font = title_font
    tcell3.alignment = Alignment(horizontal='right', vertical='center')
    style_row(ws3, 1, ncols3, fill=hdr_fill)
    style_row(ws3, 2, ncols3, fill=hdr_fill)
    add_logo(ws3)

    advertisers_ordered, adv_party, _, _ = _advertiser_order(rows2)
    weekly = defaultdict(lambda: defaultdict(float))
    platforms_by_adv = defaultdict(set)
    for r_ in rows2:
        wk = week_start_of(r_['target_date'])
        weekly[(r_['advertiser'], r_['source_platform'])][wk] += r_['spend']
        platforms_by_adv[r_['advertiser']].add(r_['source_platform'])

    dem_weekly = defaultdict(float)
    rep_weekly = defaultdict(float)
    for a in advertisers_ordered:
        p = adv_party.get(a, '')
        if p not in ('D', 'R'):
            continue
        target = dem_weekly if p == 'D' else rep_weekly
        for plat in platforms_by_adv[a]:
            for wk, v in weekly[(a, plat)].items():
                target[wk] += v

    dem_total = sum(dem_weekly.values())
    rep_total = sum(rep_weekly.values())
    np_total = party_total_spend.get('NP', 0.0)
    delta_total = dem_total - rep_total
    dem_solid, rep_solid = PARTY_SOLID['D'], PARTY_SOLID['R']

    r3 = 4
    ws3.cell(row=r3, column=1, value="DEMOCRAT TOTAL SPEND")
    ws3.cell(row=r3, column=2, value=dem_total).number_format = MONEY_TOT
    style_row(ws3, r3, ncols3, font=party_tot_font, fill=PatternFill('solid', fgColor=dem_solid))
    r3 += 1
    ws3.cell(row=r3, column=1, value="REPUBLICAN TOTAL SPEND")
    ws3.cell(row=r3, column=2, value=rep_total).number_format = MONEY_TOT
    style_row(ws3, r3, ncols3, font=party_tot_font, fill=PatternFill('solid', fgColor=rep_solid))
    r3 += 1
    ws3.cell(row=r3, column=1, value="DELTA (DEMOCRAT − REPUBLICAN)")
    ws3.cell(row=r3, column=2, value=delta_total).number_format = MONEY_TOT
    style_row(ws3, r3, ncols3, font=grand_tot_font, fill=PatternFill('solid', fgColor=NAVY), border=BORDER_GRAND)
    r3 += 2

    ws3.cell(row=r3, column=1, value="WEEK").font = hdr_font
    ws3.cell(row=r3, column=2, value="DEMOCRAT").font = hdr_font
    ws3.cell(row=r3, column=3, value="REPUBLICAN").font = hdr_font
    ws3.cell(row=r3, column=4, value="DELTA (D − R)").font = hdr_font
    style_row(ws3, r3, ncols3, fill=hdr_fill)
    r3 += 1
    first_data_row = r3
    for wk in weeks:
        d = dem_weekly.get(wk, 0.0)
        rep = rep_weekly.get(wk, 0.0)
        ws3.cell(row=r3, column=1, value=wk.strftime('%m/%d/%Y')).font = cell_font
        dcell = ws3.cell(row=r3, column=2, value=d); dcell.number_format = MONEY; dcell.font = cell_font
        rcell = ws3.cell(row=r3, column=3, value=rep); rcell.number_format = MONEY; rcell.font = cell_font
        delcell = ws3.cell(row=r3, column=4, value=d - rep); delcell.number_format = MONEY; delcell.font = cell_font
        r3 += 1
    last_data_row = r3 - 1
    r3 += 1

    ws3.conditional_formatting.add(
        f"D{first_data_row}:D{last_data_row}",
        ColorScaleRule(start_type='min', start_color=rep_solid,
                       mid_type='num', mid_value=0, mid_color='FFFFFF',
                       end_type='max', end_color=dem_solid)
    )

    ws3.column_dimensions['A'].width = 14
    ws3.column_dimensions['B'].width = 14
    ws3.column_dimensions['C'].width = 14
    ws3.column_dimensions['D'].width = 14
    ws3.freeze_panes = f"A{first_data_row}"

    ws3.cell(row=r3, column=1,
             value=f"Non-Partisan advertisers (${np_total:,.0f} total) are excluded from this Democrat-vs-Republican comparison.").font = note_font
    ws3.merge_cells(start_row=r3, start_column=1, end_row=r3, end_column=ncols3)
    r3 += 2

    add_footer(ws3, r3, ncols3)


def build(data_dir, title, output_path, today, include_dvr_tab):
    rows1 = json.load(open(os.path.join(data_dir, "raw_creative_clean.json")))
    rows2 = json.load(open(os.path.join(data_dir, "daily_spend_clean.json")))

    for r in rows1:
        r['spend'] = float(r['spend'])
        r['first_ran'] = datetime.date.fromisoformat(r['first_ran'])
        r['last_ran'] = datetime.date.fromisoformat(r['last_ran'])
    for r in rows2:
        r['spend'] = float(r['spend'])
        r['target_date'] = datetime.date.fromisoformat(r['target_date'])

    all_dates = [r['target_date'] for r in rows2] + [r['first_ran'] for r in rows1] + [r['last_ran'] for r in rows1]
    weeks = weeks_covering(all_dates, today)

    wb = Workbook()
    wb.remove(wb.active)  # tab builders each create their own sheet

    add_competitive_report_tab(wb, "Competitive Digital Report", title, weeks, rows2)
    add_creative_timeline_tab(wb, "Creative Timeline", weeks, rows1, rows2)

    if include_dvr_tab:
        _, _, _, party_total_spend = _advertiser_order(rows2)
        add_dvr_tab(wb, weeks, rows2, party_total_spend)

    wb.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--data-dir", required=True, help="Directory with raw_creative_clean.json and daily_spend_clean.json")
    parser.add_argument("--title", required=True, help="Race title shown in the report banner, e.g. PA-GOV")
    parser.add_argument("--output", required=True, help="Path to write the .xlsx report to")
    parser.add_argument("--today", default=None, help="ISO date (YYYY-MM-DD) to treat as 'today' for week bucketing; defaults to the real current date")
    parser.add_argument("--dvr-tab", action="store_true", help="Include the Democrat vs Republican tab (two-party general elections only)")
    args = parser.parse_args()

    today = datetime.date.fromisoformat(args.today) if args.today else datetime.date.today()
    out = build(args.data_dir, args.title, args.output, today, args.dvr_tab)
    print(f"saved {out}")


if __name__ == "__main__":
    main()
