"""Build a GPS Impact-branded multi-district digital competitive report --
one workbook per state/coalition, rolling up every tracked race.

Usage:
    python3 build_multi_district_report.py --data-dir DIR --title TITLE --output OUT.xlsx [--today YYYY-MM-DD]

Expects DIR to contain the same two files as build_report.py (raw_creative_clean.json,
daily_spend_clean.json), except each row additionally carries an `election` field
identifying which race it belongs to -- a 2-part statewide code (e.g. "WI-GOV",
"WI-AG") or a 3-part district code (e.g. "WI-STSEN-05", "WI-LEG-61", zero-padded
district numbers are fine, they get un-padded for display).

Produces:
    Summary            - every race rolled up by platform, GOP/DEM/delta, in
                          Statewide -> Senate -> House (or Assembly, for WI's
                          lower chamber) order
    <race> / <race> CT - one "Competitive Digital Report" tab and one "Creative
                          Timeline" tab per race, reusing the exact same styling
                          as a single-race report (see build_report.py) -- a race
                          with no tracked spend/creative yet still gets a tab,
                          it's just all zeroes rather than being skipped.
"""
import argparse
import datetime
import json
import os
import re
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_report import (  # noqa: E402
    add_competitive_report_tab, add_creative_timeline_tab, weeks_covering,
    hdr_font, hdr_fill, title_font, cell_font, footer_font, add_logo, add_footer,
    style_row, BANNER_ROW1_PT, BANNER_ROW2_PT, BODY_FONT, NAVY, MONEY, MONEY_TOT,
)

SUBTOTAL_FILL_COLOR = 'b4c6d5'
CATEGORY_ORDER = {'Statewide': 0, 'Senate': 1, 'House': 2, 'Assembly': 2, 'Other': 3}


def parse_election_code(code):
    """Splits an Ad Hawk election code into (category, sheet_name, summary_label).

    2-part codes (e.g. "WI-GOV", "WI-AG") are statewide races and are kept
    as-is. 3-part codes carry a chamber + district number: STSEN is a state
    Senate seat everywhere; LEG is a state Assembly seat in Wisconsin
    specifically (that's WI's actual name for its lower chamber) and a
    state House seat everywhere else. District numbers are un-padded
    ("05" -> "5") for the human-readable label.
    """
    parts = code.split('-')
    if len(parts) == 2:
        return 'Statewide', code, code
    if len(parts) == 3:
        state, chamber, district = parts
        try:
            dist_num = str(int(district))
        except ValueError:
            return 'Other', code, code
        if chamber == 'STSEN':
            return 'Senate', f'SD-{dist_num}', f'State Senate District {dist_num}'
        if chamber == 'LEG':
            if state == 'WI':
                return 'Assembly', f'AD-{dist_num}', f'State Assembly District {dist_num}'
            return 'House', f'HD-{dist_num}', f'State House District {dist_num}'
    return 'Other', code, code


def _race_sort_key(code):
    category, _, _ = parse_election_code(code)
    parts = code.split('-')
    if len(parts) == 3:
        try:
            return CATEGORY_ORDER[category], int(parts[2]), code
        except ValueError:
            pass
    return CATEGORY_ORDER[category], code, code


def _unique_sheet_name(name, used):
    """Excel sheet names: 31 chars max, no [ ] * ? / \\ : -- and no
    collisions (e.g. two different codes could theoretically un-pad to the
    same short name)."""
    name = re.sub(r'[\[\]*?/\\:]', '', name)[:31]
    base, n = name, 1
    while name in used:
        suffix = f' {n}'
        name = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(name)
    return name


def add_summary_tab(wb, banner_title, race_info):
    """race_info: ordered list of (summary_label, platform_stats), where
    platform_stats is the {platform: {'total','GOP','DEM'}} dict returned
    by add_competitive_report_tab for that race. Platforms are listed in a
    single order shared by every race block (by total spend across the
    whole dataset), so a race with no spend on a given platform still gets
    a $0 row instead of the block shifting around race to race."""
    platform_totals = defaultdict(float)
    for _, stats in race_info:
        for platform, s in stats.items():
            platform_totals[platform] += s['total']
    platforms_ordered = sorted(platform_totals, key=lambda p: -platform_totals[p])

    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False
    HEADERS = ["RACE", "PLATFORM", "TOTAL SPEND", "GOP TOTAL", "DEM TOTAL", "DELTA (DEM-GOP)"]
    ncols = len(HEADERS)

    ws.row_dimensions[1].height = BANNER_ROW1_PT
    ws.row_dimensions[2].height = BANNER_ROW2_PT
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=ncols)
    tcell = ws.cell(row=1, column=1, value=f"   {banner_title}")
    tcell.font = title_font
    tcell.alignment = Alignment(horizontal='right', vertical='center')
    style_row(ws, 1, ncols, fill=hdr_fill)
    style_row(ws, 2, ncols, fill=hdr_fill)
    add_logo(ws)

    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    subtotal_font = Font(name=BODY_FONT, size=9, bold=True, color=NAVY)
    subtotal_fill = PatternFill('solid', fgColor=SUBTOTAL_FILL_COLOR)
    grand_font = Font(name=BODY_FONT, size=10, bold=True, color='FFFFFF')
    grand_fill = PatternFill('solid', fgColor=NAVY)

    r = 4
    grand_total = grand_gop = grand_dem = 0.0
    for label, stats in race_info:
        block_start = r
        for platform in platforms_ordered:
            s = stats.get(platform, {'total': 0.0, 'GOP': 0.0, 'DEM': 0.0})
            ws.cell(row=r, column=2, value=platform).font = cell_font
            for col, key in ((3, 'total'), (4, 'GOP'), (5, 'DEM')):
                c = ws.cell(row=r, column=col, value=s[key])
                c.number_format = MONEY
                c.font = cell_font
            delc = ws.cell(row=r, column=6, value=s['DEM'] - s['GOP'])
            delc.number_format = MONEY
            delc.font = cell_font
            r += 1

        race_total = sum(v['total'] for v in stats.values())
        race_gop = sum(v['GOP'] for v in stats.values())
        race_dem = sum(v['DEM'] for v in stats.values())
        ws.cell(row=r, column=2, value="SUBTOTAL")
        ws.cell(row=r, column=3, value=race_total).number_format = MONEY_TOT
        ws.cell(row=r, column=4, value=race_gop).number_format = MONEY_TOT
        ws.cell(row=r, column=5, value=race_dem).number_format = MONEY_TOT
        ws.cell(row=r, column=6, value=race_dem - race_gop).number_format = MONEY_TOT
        style_row(ws, r, ncols, font=subtotal_font, fill=subtotal_fill, start=2)

        ws.merge_cells(start_row=block_start, start_column=1, end_row=r, end_column=1)
        ws.cell(row=block_start, column=1, value=label).font = Font(name=BODY_FONT, size=9, bold=True, color=NAVY)
        ws.cell(row=block_start, column=1).alignment = Alignment(vertical='top')

        grand_total += race_total
        grand_gop += race_gop
        grand_dem += race_dem
        r += 2  # blank separator row between race blocks

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.cell(row=r, column=1, value="TOTAL — ALL RACES")
    ws.cell(row=r, column=3, value=grand_total).number_format = MONEY_TOT
    ws.cell(row=r, column=4, value=grand_gop).number_format = MONEY_TOT
    ws.cell(row=r, column=5, value=grand_dem).number_format = MONEY_TOT
    ws.cell(row=r, column=6, value=grand_dem - grand_gop).number_format = MONEY_TOT
    style_row(ws, r, ncols, font=grand_font, fill=grand_fill)
    r += 2
    add_footer(ws, r, ncols)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 20
    ws.freeze_panes = "B4"


def build_multi_district(data_dir, title, output_path, today):
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

    rows2_by_race = defaultdict(list)
    for r in rows2:
        rows2_by_race[r['election']].append(r)
    rows1_by_race = defaultdict(list)
    for r in rows1:
        rows1_by_race[r['election']].append(r)

    races = sorted(set(rows2_by_race) | set(rows1_by_race), key=_race_sort_key)

    wb = Workbook()
    wb.remove(wb.active)

    used_names = set()
    race_info = []
    for code in races:
        _, short_code, label = parse_election_code(code)
        sheet_name = _unique_sheet_name(short_code, used_names)
        platform_stats = add_competitive_report_tab(
            wb, sheet_name, short_code, weeks, rows2_by_race.get(code, [])
        )
        ct_name = _unique_sheet_name(f"{short_code} CT", used_names)
        add_creative_timeline_tab(
            wb, ct_name, weeks, rows1_by_race.get(code, []), rows2_by_race.get(code, [])
        )
        race_info.append((label, platform_stats))

    add_summary_tab(wb, f"{title} DIGITAL COMPETITIVE REPORT", race_info)
    wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 1))

    wb.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--data-dir", required=True, help="Directory with raw_creative_clean.json and daily_spend_clean.json (rows must include an 'election' field)")
    parser.add_argument("--title", required=True, help="State/coalition title shown in the Summary tab banner, e.g. WI")
    parser.add_argument("--output", required=True, help="Path to write the .xlsx report to")
    parser.add_argument("--today", default=None, help="ISO date (YYYY-MM-DD) to treat as 'today' for week bucketing; defaults to the real current date")
    args = parser.parse_args()

    today = datetime.date.fromisoformat(args.today) if args.today else datetime.date.today()
    out = build_multi_district(args.data_dir, args.title, args.output, today)
    print(f"saved {out}")


if __name__ == "__main__":
    main()
