#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║           DataViz Terminal Dashboard  v2.0                       ║
║   Python · Matplotlib · Pandas · CSV Import · Excel Export       ║
╚══════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import io
import copy
import datetime

# ── dependency check ────────────────────────────────────────────────
def check_deps():
    missing = []
    for pkg in ['matplotlib', 'pandas', 'numpy', 'openpyxl']:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"\n[ERROR] Missing packages: {', '.join(missing)}")
        print(f"  Run:  pip install {' '.join(missing)}\n")
        sys.exit(1)

check_deps()

import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# ── colour palette ───────────────────────────────────────────────────
BLUE   = '#58a6ff'
GREEN  = '#3fb950'
PURPLE = '#d2a8ff'
ORANGE = '#ffa657'
RED    = '#f78166'
CYAN   = '#39d353'
GOLD   = '#e3b341'
PINK   = '#f0883e'
COLORS = [BLUE, GREEN, PURPLE, ORANGE, RED, CYAN, GOLD, PINK]

BG          = '#0d1117'
BG2         = '#161b22'
BG3         = '#1c2230'
BORDER_CLR  = '#30363d'
TEXT        = '#e6edf3'
TEXT2       = '#8b949e'
TEXT3       = '#6e7681'

# ╔══════════════════════════════════════════════════════════════════╗
# ║  TERMINAL HELPERS                                               ║
# ╚══════════════════════════════════════════════════════════════════╝

def cls():
    os.system('cls' if os.name == 'nt' else 'clear')

def cprint(text, color='', bold=False, end='\n'):
    codes = {
        'blue':   '\033[94m', 'green':  '\033[92m', 'yellow': '\033[93m',
        'red':    '\033[91m', 'cyan':   '\033[96m', 'purple': '\033[95m',
        'white':  '\033[97m', 'gray':   '\033[90m', '':       ''
    }
    bold_code = '\033[1m' if bold else ''
    reset = '\033[0m'
    print(f"{bold_code}{codes.get(color, '')}{text}{reset}", end=end)

def divider(char='─', width=66, color='gray'):
    cprint(char * width, color)

def prompt(msg, color='cyan'):
    codes = {'cyan': '\033[96m', 'green': '\033[92m', 'yellow': '\033[93m'}
    return input(f"{codes.get(color, '')}{msg}\033[0m")

def header():
    cls()
    print()
    cprint('  ╔══════════════════════════════════════════════════════════════╗', 'blue')
    cprint('  ║                                                              ║', 'blue')
    cprint('  ║', 'blue', end='')
    print('\033[1m\033[96m  DataViz Terminal Dashboard  v2.0\033[0m\033[94m                       ║')
    cprint('  ║   Python · Matplotlib · Pandas · CSV · Excel Export          ║', 'blue')
    cprint('  ║                                                              ║', 'blue')
    cprint('  ╚══════════════════════════════════════════════════════════════╝', 'blue')
    print()

def print_df(df, max_rows=8):
    divider()
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 100)
    pd.set_option('display.float_format', '{:.2f}'.format)
    cprint(df.head(max_rows).to_string(index=False), 'white')
    divider()

# ╔══════════════════════════════════════════════════════════════════╗
# ║  MATPLOTLIB STYLE                                               ║
# ╚══════════════════════════════════════════════════════════════════╝

def setup_style():
    plt.rcParams.update({
        'figure.facecolor':  BG,
        'axes.facecolor':    BG2,
        'axes.edgecolor':    BORDER_CLR,
        'axes.labelcolor':   TEXT2,
        'axes.titlecolor':   TEXT,
        'axes.titlesize':    11,
        'axes.titleweight':  'bold',
        'axes.titlepad':     12,
        'axes.grid':         True,
        'grid.color':        BORDER_CLR,
        'grid.linewidth':    0.5,
        'grid.alpha':        0.6,
        'xtick.color':       TEXT3,
        'ytick.color':       TEXT3,
        'xtick.labelsize':   9,
        'ytick.labelsize':   9,
        'text.color':        TEXT,
        'font.family':       'monospace',
        'legend.facecolor':  BG3,
        'legend.edgecolor':  BORDER_CLR,
        'legend.labelcolor': TEXT2,
        'legend.fontsize':   9,
        'figure.titlesize':  13,
        'figure.titleweight':'bold',
    })

def add_watermark(fig, label):
    fig.text(0.99, 0.01, f'dataviz.py v2.0 · {label}',
             ha='right', va='bottom', fontsize=8,
             color=TEXT3, style='italic', family='monospace')

# ╔══════════════════════════════════════════════════════════════════╗
# ║  INDIVIDUAL CHART BUILDERS                                      ║
# ╚══════════════════════════════════════════════════════════════════╝

def _ax_style(ax):
    ax.set_facecolor(BG2)
    ax.tick_params(colors=TEXT3)
    ax.grid(True, color=BORDER_CLR, alpha=0.5)
    for spine in ax.spines.values():
        spine.set_edgecolor(BORDER_CLR)

def make_line_fig(months, series, title_suffix=''):
    fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
    _ax_style(ax)
    ax.set_title(f'plt.plot()  ·  {title_suffix}', color=BLUE, pad=10)
    for i, (name, vals) in enumerate(series.items()):
        c = COLORS[i % len(COLORS)]
        ax.plot(months, vals, color=c, linewidth=2.2,
                marker='o', markersize=5, label=name)
        ax.fill_between(months, vals, alpha=0.08, color=c)
    ax.legend()
    fig.tight_layout()
    return fig

def make_bar_fig(categories, values, ylabel, title_suffix=''):
    fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
    _ax_style(ax)
    ax.set_title(f'plt.bar()  ·  {title_suffix}', color=GREEN, pad=10)
    bar_colors = (COLORS * 4)[:len(categories)]
    bars = ax.bar(categories, values, color=bar_colors,
                  edgecolor=BG, linewidth=0.8, width=0.6)
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2,
                bar.get_height() * 1.01,
                f'{val:.1f}', ha='center', va='bottom',
                fontsize=8, color=TEXT2, fontfamily='monospace')
    ax.set_ylabel(ylabel, color=TEXT2)
    ax.tick_params(axis='x', rotation=20)
    ax.grid(True, color=BORDER_CLR, alpha=0.5, axis='y')
    fig.tight_layout()
    return fig

def make_pie_fig(labels, values, title_suffix=''):
    fig, ax = plt.subplots(figsize=(8, 6), facecolor=BG)
    ax.set_facecolor(BG2)
    ax.set_title(f'plt.pie()  ·  {title_suffix}', color=PURPLE, pad=10)
    pie_colors = (COLORS * 4)[:len(labels)]
    wedges, texts, autotexts = ax.pie(
        values, labels=labels, autopct='%1.1f%%',
        colors=pie_colors, startangle=90,
        wedgeprops=dict(width=0.55, edgecolor=BG, linewidth=2),
        textprops=dict(color=TEXT2, fontsize=8, fontfamily='monospace')
    )
    for at in autotexts:
        at.set_color(BG)
        at.set_fontweight('bold')
    fig.tight_layout()
    return fig

def make_scatter_fig(x, y, xlabel, ylabel, title_suffix=''):
    fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
    _ax_style(ax)
    ax.set_title(f'plt.scatter()  ·  {title_suffix}', color=ORANGE, pad=10)
    ax.scatter(x, y, color=ORANGE, alpha=0.78,
               edgecolors=BG2, linewidth=0.8, s=60, zorder=3)
    if len(x) > 2:
        try:
            z = np.polyfit(x, y, 1)
            p = np.poly1d(z)
            xr = np.linspace(min(x), max(x), 100)
            ax.plot(xr, p(xr), color=BLUE, linewidth=1.5,
                    linestyle='--', alpha=0.7)
        except Exception:
            pass
    ax.set_xlabel(xlabel, color=TEXT2)
    ax.set_ylabel(ylabel, color=TEXT2)
    fig.tight_layout()
    return fig

def make_barh_fig(categories, values, xlabel, title_suffix=''):
    fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
    _ax_style(ax)
    ax.set_title(f'plt.barh()  ·  {title_suffix}', color=CYAN, pad=10)
    pairs = sorted(zip(values, categories))
    sv, sc = zip(*pairs)
    ax.barh(list(sc), list(sv), color=CYAN, alpha=0.82,
            edgecolor=BG, height=0.6)
    ax.set_xlabel(xlabel, color=TEXT2)
    ax.tick_params(axis='y', labelsize=8)
    ax.grid(True, color=BORDER_CLR, alpha=0.5, axis='x')
    fig.tight_layout()
    return fig

# ╔══════════════════════════════════════════════════════════════════╗
# ║  EXCEL EXPORT ENGINE                                            ║
# ╚══════════════════════════════════════════════════════════════════╝

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=130, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf

def _thin_border():
    s = Side(style='thin', color='30363D')
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr_cell(cell, text=None, bg='1F2D3D', fg='E6EDF3'):
    if text is not None:
        cell.value = text
    cell.font = Font(bold=True, color=fg, name='Consolas', size=10)
    cell.fill = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = _thin_border()

def _data_cell(cell, val, even=True):
    cell.value = val
    cell.fill = PatternFill('solid', start_color='161B22' if even else '1C2230')
    cell.font = Font(color='C9D1D9', name='Consolas', size=9)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = _thin_border()

def write_df_to_sheet(ws, df, start_row=1, start_col=1):
    for ci, col in enumerate(df.columns, start_col):
        _hdr_cell(ws.cell(row=start_row, column=ci), str(col))
        ws.column_dimensions[get_column_letter(ci)].width = max(14, len(str(col)) + 4)
    for ri, row in enumerate(df.itertuples(index=False), start_row + 1):
        for ci, val in enumerate(row, start_col):
            _data_cell(ws.cell(row=ri, column=ci), val, ri % 2 == 0)
    ws.row_dimensions[start_row].height = 22

def embed_charts_excel(figs_with_labels, df, title, filename):
    wb = Workbook()

    # ── Sheet 1: Data ────────────────────────────────────────────────
    ws_data = wb.active
    ws_data.title = 'Data'
    ws_data.sheet_view.showGridLines = False
    ws_data.sheet_properties.tabColor = '58A6FF'

    ws_data.merge_cells('A1:H1')
    ws_data['A1'].value = f'  {title}  ·  dataviz.py v2.0'
    ws_data['A1'].font = Font(bold=True, color='58A6FF', name='Consolas', size=13)
    ws_data['A1'].fill = PatternFill('solid', start_color='0D1117')
    ws_data['A1'].alignment = Alignment(horizontal='left', vertical='center')
    ws_data.row_dimensions[1].height = 30

    ws_data.merge_cells('A2:H2')
    ws_data['A2'].value = (f'  Exported: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}'
                           f'  ·  Rows: {len(df)}  ·  Columns: {len(df.columns)}')
    ws_data['A2'].font = Font(color='6E7681', name='Consolas', size=9, italic=True)
    ws_data['A2'].fill = PatternFill('solid', start_color='0D1117')
    ws_data.row_dimensions[2].height = 18

    write_df_to_sheet(ws_data, df, start_row=4, start_col=1)
    ws_data.freeze_panes = 'A5'

    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if num_cols:
        stats_df = df[num_cols].describe().round(2).reset_index()
        stats_df.rename(columns={'index': 'Stat'}, inplace=True)
        sr = len(df) + 7
        ws_data.merge_cells(f'A{sr}:H{sr}')
        ws_data[f'A{sr}'].value = '  STATISTICS  (df.describe())'
        ws_data[f'A{sr}'].font = Font(bold=True, color='3FB950', name='Consolas', size=10)
        ws_data[f'A{sr}'].fill = PatternFill('solid', start_color='0D1117')
        ws_data.row_dimensions[sr].height = 22
        write_df_to_sheet(ws_data, stats_df, start_row=sr + 1, start_col=1)

    # ── Sheet 2: Charts ──────────────────────────────────────────────
    ws_charts = wb.create_sheet('Charts')
    ws_charts.sheet_view.showGridLines = False
    ws_charts.sheet_properties.tabColor = 'D2A8FF'

    for r in range(1, 200):
        for c in range(1, 30):
            cell = ws_charts.cell(row=r, column=c)
            cell.fill = PatternFill('solid', start_color='0D1117')

    ws_charts.merge_cells('A1:P1')
    ws_charts['A1'].value = f'  {title}  ·  Charts Dashboard'
    ws_charts['A1'].font = Font(bold=True, color='58A6FF', name='Consolas', size=13)
    ws_charts['A1'].fill = PatternFill('solid', start_color='0D1117')
    ws_charts['A1'].alignment = Alignment(horizontal='left', vertical='center')
    ws_charts.row_dimensions[1].height = 30

    ws_charts.merge_cells('A2:P2')
    ts = f'matplotlib v{matplotlib.__version__}'
    ws_charts['A2'].value = f'  plt.plot() · plt.bar() · plt.scatter() · plt.pie()  ·  {ts}'
    ws_charts['A2'].font = Font(color='6E7681', name='Consolas', size=9, italic=True)
    ws_charts['A2'].fill = PatternFill('solid', start_color='0D1117')
    ws_charts.row_dimensions[2].height = 18

    COL_POSITIONS = [2, 12]
    ROW_START = 4
    ROW_STEP  = 34

    for idx, (fig, label) in enumerate(figs_with_labels):
        buf = fig_to_bytes(fig)
        img = XLImage(buf)
        img.width  = 620
        img.height = 400
        col_idx = COL_POSITIONS[idx % 2]
        row_pos = ROW_START + (idx // 2) * ROW_STEP
        col_ltr = get_column_letter(col_idx)

        lbl_cell = ws_charts.cell(row=row_pos - 1, column=col_idx)
        lbl_cell.value = f'  {label}'
        lbl_cell.font = Font(bold=True, color='8B949E', name='Consolas', size=9)
        lbl_cell.fill = PatternFill('solid', start_color='0D1117')

        ws_charts.add_image(img, f'{col_ltr}{row_pos}')

    # ── Sheet 3: README ──────────────────────────────────────────────
    ws_info = wb.create_sheet('README')
    ws_info.sheet_view.showGridLines = False
    ws_info.sheet_properties.tabColor = '3FB950'

    lines = [
        ('', ''),
        ('  DataViz Terminal Dashboard  v2.0', '58A6FF'),
        ('', ''),
        ('  HOW THIS FILE WAS GENERATED', '3FB950'),
        ('  ──────────────────────────────────────────────────────', '30363D'),
        (f'  Tool       :  dataviz.py v2.0', 'C9D1D9'),
        (f'  Title      :  {title}', 'C9D1D9'),
        (f'  Exported   :  {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 'C9D1D9'),
        (f'  Rows       :  {len(df)}', 'C9D1D9'),
        (f'  Columns    :  {len(df.columns)}', 'C9D1D9'),
        ('', ''),
        ('  SHEETS', '3FB950'),
        ('  ──────────────────────────────────────────────────────', '30363D'),
        ('  Data       :  Raw data + descriptive statistics', 'C9D1D9'),
        ('  Charts     :  All matplotlib charts embedded as images', 'C9D1D9'),
        ('  README     :  This page', 'C9D1D9'),
        ('', ''),
        ('  LIBRARIES', '3FB950'),
        ('  ──────────────────────────────────────────────────────', '30363D'),
        ('  matplotlib :  Chart rendering (line, bar, scatter, pie)', 'C9D1D9'),
        ('  pandas     :  Data loading, manipulation, statistics', 'C9D1D9'),
        ('  numpy      :  Numerical operations', 'C9D1D9'),
        ('  openpyxl   :  Excel creation and image embedding', 'C9D1D9'),
    ]

    for ri, (text, color) in enumerate(lines, 1):
        cell = ws_info.cell(row=ri, column=1, value=text)
        is_big = 'DataViz' in text
        is_hdr = any(k in text for k in ['HOW', 'SHEETS', 'LIBRARIES'])
        cell.font = Font(color=color or '0D1117', name='Consolas',
                         size=13 if is_big else 10,
                         bold=is_big or is_hdr)
        cell.fill = PatternFill('solid', start_color='0D1117')
        ws_info.row_dimensions[ri].height = 20

    ws_info.column_dimensions['A'].width = 65

    wb.save(filename)
    return filename

# ╔══════════════════════════════════════════════════════════════════╗
# ║  EXPORT HELPER                                                  ║
# ╚══════════════════════════════════════════════════════════════════╝

def ask_export(figs, df, title):
    print()
    ans = prompt('  → Export to Excel (.xlsx) with charts embedded? [Y/n]: ').strip().lower()
    if ans in ('', 'y', 'yes'):
        default = title.lower().replace(' ', '_') + '.xlsx'
        fname = prompt(f'  → Filename [{default}]: ').strip()
        if not fname:
            fname = default
        if not fname.endswith('.xlsx'):
            fname += '.xlsx'
        cprint('\n  ⟳  Building Excel workbook...', 'cyan')
        out = embed_charts_excel(figs, df, title, fname)
        cprint(f'  ✓  Saved → {os.path.abspath(out)}', 'green', bold=True)
        cprint('     Sheets: Data  ·  Charts  ·  README', 'gray')
    print()
    input('  Press Enter to return to menu...')
    plt.close('all')

# ╔══════════════════════════════════════════════════════════════════╗
# ║  PRE-LOADED DATASETS                                            ║
# ╚══════════════════════════════════════════════════════════════════╝

def build_students():
    np.random.seed(42)
    names = [
        'Arjun S.','Priya M.','Rahul K.','Sneha T.','Dev R.',
        'Meera J.','Kunal P.','Ananya V.','Rohan B.','Pooja N.',
        'Vikram S.','Divya K.','Amit T.','Riya M.','Saurabh P.'
    ]
    subjects = ['Math','Science','English','History','Art']
    data = {s: np.random.randint(48, 99, len(names)) for s in subjects}
    df = pd.DataFrame(data)
    df.insert(0, 'Student', names)
    df['Average'] = df[subjects].mean(axis=1).round(1)
    df['Grade'] = df['Average'].apply(
        lambda x: 'A' if x >= 85 else ('B' if x >= 70 else ('C' if x >= 55 else 'D'))
    )
    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug']
    trends = {
        'Math':    [72,75,70,78,82,80,85,88],
        'Science': [65,68,71,74,70,76,79,83],
        'English': [80,78,82,79,83,85,81,87],
    }
    gc = df['Grade'].value_counts()
    return {
        'type': 'students', 'df': df, 'subjects': subjects,
        'months': months, 'trends': trends,
        'grade_labels': list(gc.index), 'grade_vals': list(gc.values),
        'title': 'Student Performance Dashboard'
    }

def build_sales():
    months   = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug']
    revenue  = {
        'Electronics':    [320,290,380,410,350,430,480,510],
        'Clothing':       [180,210,170,195,230,215,240,260],
        'Home & Kitchen': [90, 85, 110,120,100,130,115,140],
    }
    regions  = ['North','South','East','West','Central','Online']
    reg_sale = [820,650,490,710,380,1050]
    cat_sh   = {'Electronics':42,'Clothing':28,'Home & Kitchen':15,'Books':9,'Sports':6}
    products = ['Smartphone','Earbuds','Kurta Set','Smart Watch',
                'Pressure Cooker','Yoga Mat','Backpack','Novel']
    units    = [482,316,580,201,340,410,228,612]
    prices   = [2000,1000,400,3000,500,300,500,120]
    df = pd.DataFrame({
        'Product':     products,
        'Units':       units,
        'Price (Rs)':  prices,
        'Revenue (Rs)':[u*p for u,p in zip(units,prices)]
    })
    return {
        'type': 'sales', 'df': df, 'months': months,
        'revenue': revenue, 'regions': regions, 'reg_sales': reg_sale,
        'cat_share': cat_sh, 'title': 'Sales Analytics Dashboard'
    }

def run_preloaded(data):
    setup_style()
    title = data['title']
    figs  = []

    if data['type'] == 'students':
        subjs = data['subjects']
        avgs  = [data['df'][s].mean() for s in subjs]
        figs  = [
            (make_line_fig(data['months'], data['trends'], 'Score Trends by Subject'),
             'Score Trends — Line'),
            (make_bar_fig(subjs, avgs, 'Avg Score', 'Average Score per Subject'),
             'Subject Averages — Bar'),
            (make_pie_fig(data['grade_labels'], data['grade_vals'], 'Grade Distribution'),
             'Grade Distribution — Pie'),
            (make_barh_fig(list(data['df']['Student']),
                           list(data['df']['Average']),
                           'Average Score', 'Student Rankings'),
             'Student Rankings — Barh'),
            (make_scatter_fig(list(range(1, len(data['df'])+1)),
                              list(data['df']['Average']),
                              'Student Index', 'Average Score', 'Score Distribution'),
             'Score Distribution — Scatter'),
        ]

    elif data['type'] == 'sales':
        figs = [
            (make_line_fig(data['months'], data['revenue'], 'Monthly Revenue by Category'),
             'Revenue Trends — Line'),
            (make_bar_fig(data['regions'], data['reg_sales'], 'Revenue (K)', 'Sales by Region'),
             'Regional Sales — Bar'),
            (make_pie_fig(list(data['cat_share'].keys()),
                          list(data['cat_share'].values()), 'Revenue by Category'),
             'Category Share — Pie'),
            (make_barh_fig(list(data['df']['Product']),
                           list(data['df']['Revenue (Rs)']),
                           'Revenue (Rs)', 'Revenue by Product'),
             'Product Revenue — Barh'),
            (make_scatter_fig(list(data['df']['Units']),
                              list(data['df']['Revenue (Rs)']),
                              'Units Sold', 'Revenue (Rs)', 'Units vs Revenue'),
             'Units vs Revenue — Scatter'),
        ]

    cprint(f'\n  ⟳  Opening {len(figs)} chart windows...', 'cyan')
    for fig, lbl in figs:
        add_watermark(fig, lbl)
        plt.figure(fig.number)
        plt.show(block=False)

    ask_export(figs, data['df'], title)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  CSV IMPORT                                                     ║
# ╚══════════════════════════════════════════════════════════════════╝

def csv_import():
    header()
    cprint('  ┌─ CSV IMPORT MODE ──────────────────────────────────────────┐', 'yellow')
    cprint('  │  Paste the full path to your .csv file below               │', 'yellow')
    cprint('  └─────────────────────────────────────────────────────────────┘', 'yellow')
    print()

    while True:
        raw = prompt('  → CSV file path: ').strip().strip('"').strip("'")
        if not raw:
            cprint('  [!] No path entered.', 'red')
            continue
        if not os.path.isfile(raw):
            cprint(f'  [!] File not found: {raw}', 'red')
            continue
        break

    cprint(f'\n  ⟳  Reading {os.path.basename(raw)} ...', 'cyan')
    try:
        df = pd.read_csv(raw)
    except Exception as e:
        cprint(f'  [ERROR] {e}', 'red')
        input('  Press Enter...'); return

    cprint(f'  ✓  Loaded {len(df)} rows × {len(df.columns)} columns', 'green')
    print()
    print_df(df)

    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    cat_cols = df.select_dtypes(exclude=[np.number]).columns.tolist()

    if not num_cols:
        cprint('  [!] No numeric columns — cannot plot.', 'red')
        input('  Press Enter...'); return

    cprint(f'  Numeric : {", ".join(num_cols)}', 'green')
    cprint(f'  Text    : {", ".join(cat_cols) if cat_cols else "none"}', 'gray')
    print()

    label_col = cat_cols[0] if cat_cols else None
    title = os.path.splitext(os.path.basename(raw))[0].replace('_', ' ').title()

    go = prompt('  → Generate charts for all numeric columns? [Y/n]: ').strip().lower()
    if go not in ('', 'y', 'yes'):
        input('  Aborted. Press Enter...'); return

    setup_style()
    figs = []

    for col in num_cols:
        vals = df[col].dropna().tolist()
        cats = (df[label_col].astype(str).tolist()
                if label_col else [str(i+1) for i in range(len(vals))])
        cats = cats[:len(vals)]

        figs.append((make_bar_fig(cats[:20], vals[:20], col,
                                  f'{col} by {label_col or "Index"}'),
                     f'{col} — Bar'))
        figs.append((make_barh_fig(cats[:15], vals[:15], col,
                                   f'{col} Ranked'),
                     f'{col} — Barh'))
        figs.append((make_line_fig([str(c)[:10] for c in cats[:20]],
                                   {col: vals[:20]}, f'{col} Trend'),
                     f'{col} — Line'))

    if label_col and len(df[num_cols[0]]) <= 12:
        figs.append((make_pie_fig(df[label_col].astype(str).tolist(),
                                  df[num_cols[0]].abs().tolist(),
                                  f'{num_cols[0]} Distribution'),
                     f'{num_cols[0]} — Pie'))

    if len(num_cols) >= 2:
        xv = df[num_cols[0]].dropna().tolist()
        yv = df[num_cols[1]].dropna().tolist()
        ml = min(len(xv), len(yv))
        figs.append((make_scatter_fig(xv[:ml], yv[:ml],
                                      num_cols[0], num_cols[1],
                                      f'{num_cols[0]} vs {num_cols[1]}'),
                     f'{num_cols[0]} vs {num_cols[1]} — Scatter'))

    cprint(f'\n  ⟳  Opening {len(figs)} chart windows...', 'cyan')
    for fig, lbl in figs:
        add_watermark(fig, lbl)
        plt.figure(fig.number)
        plt.show(block=False)

    ask_export(figs, df, title)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  MANUAL ENTRY                                                   ║
# ╚══════════════════════════════════════════════════════════════════╝

def manual_entry():
    header()
    cprint('  ┌─ MANUAL DATA ENTRY MODE ──────────────────────────────────┐', 'cyan')
    cprint('  │  Type category names and values directly in the terminal  │', 'cyan')
    cprint('  └───────────────────────────────────────────────────────────┘', 'cyan')
    print()

    title = prompt('  → Dashboard title: ').strip() or 'Custom Dashboard'
    label = prompt('  → Y-axis label (e.g. Score, Amount, Units): ').strip() or 'Value'
    print()

    while True:
        try:
            n = int(prompt('  → Number of categories (2-20): '))
            if 2 <= n <= 20:
                break
            cprint('  [!] Enter a number between 2 and 20.', 'red')
        except ValueError:
            cprint('  [!] Enter a whole number.', 'red')

    print()
    cprint(f'  Enter {n} entries:', 'yellow')
    divider()

    categories, values = [], []
    for i in range(1, n + 1):
        while True:
            name = prompt(f'  [{i}/{n}] Category name : ').strip()
            if name:
                break
            cprint('       [!] Cannot be empty.', 'red')
        while True:
            try:
                val = float(prompt(f'  [{i}/{n}] {label:<14}: '))
                break
            except ValueError:
                cprint('       [!] Enter a valid number.', 'red')
        categories.append(name)
        values.append(val)
        cprint(f'       ✓  {name} = {val}', 'green')

    df = pd.DataFrame({'Category': categories, label: values})
    df['% Share'] = (df[label].abs() / df[label].abs().sum() * 100).round(1)

    print()
    divider()
    cprint(f'  PREVIEW  ·  {title}', 'purple', bold=True)
    print_df(df, max_rows=25)
    cprint(
        f'  Min: {min(values):.2f}  │  Max: {max(values):.2f}  │  '
        f'Mean: {sum(values)/len(values):.2f}  │  Total: {sum(values):.2f}',
        'gray'
    )
    divider()
    print()

    go = prompt('  → Generate charts? [Y/n]: ').strip().lower()
    if go not in ('', 'y', 'yes'):
        cprint('\n  Aborted.\n', 'gray')
        return

    setup_style()
    abs_vals = [abs(v) for v in values]
    figs = [
        (make_bar_fig(categories, values, label, f'{label} Comparison'),     'Bar Chart'),
        (make_barh_fig(categories, values, label, f'{label} Ranked'),         'Barh Chart'),
        (make_pie_fig(categories, abs_vals, 'Distribution'),                  'Pie Chart'),
        (make_line_fig(categories, {label: values}, 'Trend by Category'),     'Line Chart'),
    ]

    cprint('\n  ⟳  Opening chart windows...', 'cyan')
    for fig, lbl in figs:
        add_watermark(fig, lbl)
        plt.figure(fig.number)
        plt.show(block=False)

    ask_export(figs, df, title)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  MENUS                                                          ║
# ╚══════════════════════════════════════════════════════════════════╝

PRELOADED = {
    '1': ('Student Performance', '42 students · 6 subjects · grades', build_students),
    '2': ('Sales Analytics',     'Revenue · regions · category share', build_sales),
}

def preloaded_menu():
    header()
    cprint('  ┌─ PRE-LOADED DATASETS ──────────────────────────────────────┐', 'blue')
    for k, (name, desc, _) in PRELOADED.items():
        cprint(f'  │  [{k}]  {name:<24} {desc:<34}│', 'white')
    cprint('  │  [0]  ← Back                                               │', 'gray')
    cprint('  └─────────────────────────────────────────────────────────────┘', 'blue')
    print()

    while True:
        choice = prompt('  → Select dataset: ').strip()
        if choice == '0':
            return
        if choice in PRELOADED:
            name, _, builder = PRELOADED[choice]
            cprint(f'\n  ⟳  Loading "{name}"...', 'cyan')
            data = builder()
            print()
            cprint(f'  DATA PREVIEW  ·  {name}', 'purple', bold=True)
            print_df(data['df'])
            go = prompt('  → Open charts? [Y/n]: ').strip().lower()
            if go in ('', 'y', 'yes'):
                run_preloaded(data)
            return
        cprint('  [!] Enter 1, 2, or 0.', 'red')

def main_menu():
    while True:
        header()
        cprint('  ┌─ MAIN MENU ────────────────────────────────────────────────┐', 'green')
        cprint('  │                                                             │', 'green')
        cprint('  │   [1]  Pre-loaded datasets   Students / Sales              │', 'white')
        cprint('  │   [2]  Manual data entry     Type your own data            │', 'white')
        cprint('  │   [3]  Import CSV            Auto-detect & plot all        │', 'white')
        cprint('  │   [q]  Quit                                                 │', 'gray')
        cprint('  │                                                             │', 'green')
        cprint('  └─────────────────────────────────────────────────────────────┘', 'green')
        print()
        cprint('  Charts : line · bar · barh · scatter · pie', 'gray')
        cprint('  Export : Excel (.xlsx) — embedded charts + data + stats', 'gray')
        print()

        choice = prompt('  → Select option: ').strip().lower()
        if choice == '1':
            preloaded_menu()
        elif choice == '2':
            manual_entry()
        elif choice == '3':
            csv_import()
        elif choice in ('q', 'quit', 'exit'):
            print()
            cprint('  Goodbye! Happy plotting! 🐍', 'cyan', bold=True)
            print()
            sys.exit(0)
        else:
            cprint('\n  [!] Invalid. Enter 1, 2, 3, or q.\n', 'red')
            input('  Press Enter to continue...')

if __name__ == '__main__':
    try:
        main_menu()
    except Exception as e:
        print(f"\n[ERROR] {e}")
    input("\nPress Enter to exit...")
