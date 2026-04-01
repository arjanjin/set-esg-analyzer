# ============================================================
# AI-Powered ESG Screening System for SET50
# KBS KMITL — Senior Project
# Tech: Python + Claude API + yfinance + CFA Framework
# ============================================================

# Cell 1 — Install
# !pip install yfinance anthropic openpyxl matplotlib -q

# Cell 2 — Thai Font
# import urllib.request, os, matplotlib
# from matplotlib import font_manager
# os.makedirs('/usr/share/fonts/truetype/custom', exist_ok=True)
# url = "https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Regular.ttf"
# font_path = "/usr/share/fonts/truetype/custom/Sarabun-Regular.ttf"
# urllib.request.urlretrieve(url, font_path)
# font_manager.fontManager.addfont(font_path)
# prop = font_manager.FontProperties(fname=font_path)
# matplotlib.rcParams['font.family'] = prop.get_name()
# matplotlib.rcParams['axes.unicode_minus'] = False

import yfinance as yf
import pandas as pd
import anthropic
import json
import re
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from google.colab import userdata

client = anthropic.Anthropic(api_key=userdata.get('ANTHROPIC_API_KEY'))


# ============================================================
# STEP 1 — Get Financial Data
# ============================================================

def get_financials(ticker_symbol):
    t = yf.Ticker(ticker_symbol)
    info = t.info
    fin = t.financials

    revenue = None
    for k in ['Total Revenue', 'Net Interest Income', 'Interest Income']:
        if k in fin.index:
            revenue = fin.loc[k].iloc[0]
            break
    if not revenue or revenue == 0:
        revenue = 1

    gross = 0
    for k in ['Gross Profit', 'Net Interest Income']:
        if k in fin.index:
            gross = fin.loc[k].iloc[0]
            break

    net = 0
    for k in ['Net Income', 'Net Income Common Stockholders']:
        if k in fin.index:
            net = fin.loc[k].iloc[0]
            break

    return {
        'info': info,
        'revenue': revenue,
        'gross_margin': round(gross / revenue * 100, 2),
        'net_margin': round(net / revenue * 100, 2),
        'roe': round(info.get('returnOnEquity', 0) * 100, 2),
        'pe': info.get('trailingPE', 'N/A'),
        'de': info.get('debtToEquity', 'N/A'),
        'sector': info.get('sector', 'N/A')
    }


# ============================================================
# STEP 2 — ESG Analysis with Claude API
# ============================================================

def parse_json_safe(text):
    text = text.strip()
    text = text.replace('```json', '').replace('```', '').strip()
    return json.loads(text)

def analyze_esg(ticker_symbol, d):
    prompt = f"""คุณเป็น ESG Analyst ตลาดหุ้นไทย

{ticker_symbol}:
Revenue {d['revenue']/1e9:.1f}B Gross {d['gross_margin']}% Net {d['net_margin']}% ROE {d['roe']}% PE {d['pe']} DE {d['de']} Sector {d['sector']}

ให้คะแนน esg_e esg_s esg_g ระหว่าง 0-100 ตาม Sector จริง
esg_total คือค่าเฉลี่ยของทั้งสาม

ตอบ JSON เท่านั้น ไม่มี markdown:
{{"name":"ชื่อไทย","esg_e":65,"esg_s":70,"esg_g":68,"esg_total":68,"recommendation":"HOLD","reason_e":"x","reason_s":"x","reason_g":"x","summary":"x"}}"""

    msg = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}]
    )
    result = parse_json_safe(msg.content[0].text)
    result.update({
        'ticker': ticker_symbol,
        'gross_margin': d['gross_margin'],
        'net_margin': d['net_margin'],
        'roe': d['roe'],
        'pe': d['pe'],
        'sector': d['sector']
    })
    return result


# ============================================================
# STEP 3 — DCF Valuation with Claude API
# ============================================================

def dcf_valuation(row):
    de = row.get('de', 'N/A')
    if str(de) == 'nan':
        t = yf.Ticker(row['ticker'])
        de = t.info.get('debtToEquity', 'N/A')

    prompt = f"""คุณเป็น CFA Analyst ทำ DCF Valuation ตลาดหุ้นไทย

{row['ticker']}:
Sector {row['sector']} ROE {row['roe']}% Net Margin {row['net_margin']}% PE {row['pe']} DE {de}

สมมติฐาน: WACC 9% Terminal Growth 3% Forecast 5 ปี
Growth Rate ตาม Sector: Energy 4-6% Banking 5-7% Retail 6-8% Tech/Healthcare 7-10%

ตอบ JSON เท่านั้น ไม่มี markdown:
{{"wacc":9.0,"growth_rate":6.5,"terminal_growth":3.0,"valuation":"Undervalued/Overvalued/Fairly Valued","upside_downside":"+xx%","dcf_summary":"สรุป 1 ประโยค"}}"""

    msg = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=500,
        messages=[{"role": "user", "content": prompt}]
    )
    return parse_json_safe(msg.content[0].text)


# ============================================================
# STEP 4 — Run ESG Analysis
# ============================================================

tickers = [
    "CPALL.BK", "BDMS.BK",
    "PTT.BK",   "GULF.BK",
    "SCB.BK",   "KBANK.BK",
    "AOT.BK",   "ADVANC.BK",
    "SCC.BK",   "DELTA.BK"
]

results = []
print("ESG Analysis SET50")
print("=" * 60)

for t in tickers:
    print(f"วิเคราะห์ {t}...")
    try:
        d = get_financials(t)
        r = analyze_esg(t, d)
        results.append(r)
        print(f"  {r['name']}  E:{r['esg_e']} S:{r['esg_s']} G:{r['esg_g']} รวม:{r['esg_total']} | {r['recommendation']}")
    except Exception as e:
        print(f"  Error: {e}")

df = pd.DataFrame(results)
df_sorted = df.sort_values('esg_total', ascending=False).reset_index(drop=True)


# ============================================================
# STEP 5 — Run DCF Valuation
# ============================================================

print("\nDCF Valuation Analysis")
print("=" * 60)

dcf_results = []
for _, row in df_sorted.iterrows():
    try:
        dcf = dcf_valuation(row)
        dcf_results.append({
            'ticker': row['ticker'],
            'name': row['name'],
            'esg_total': row['esg_total'],
            'recommendation_esg': row['recommendation'],
            'wacc': dcf['wacc'],
            'growth_rate': dcf['growth_rate'],
            'valuation': dcf['valuation'],
            'upside': dcf['upside_downside'],
            'dcf_summary': dcf['dcf_summary'],
            'sector': row['sector'],
            'roe': row['roe'],
            'net_margin': row['net_margin'],
            'pe': row['pe']
        })
        print(f"\n{row['ticker']} -- {row['name']}")
        print(f"  Valuation : {dcf['valuation']}")
        print(f"  Upside    : {dcf['upside_downside']}")
    except Exception as e:
        print(f"\n{row['ticker']} Error: {e}")

df_dcf = pd.DataFrame(dcf_results)


# ============================================================
# STEP 6 — Charts
# ============================================================

colors_map = {'BUY': '#2ecc71', 'HOLD': '#f39c12', 'AVOID': '#e74c3c'}

# Chart 1 — ESG Ranking
fig, ax = plt.subplots(figsize=(12, 6))
bar_colors = [colors_map.get(r, '#3498db') for r in df_sorted['recommendation']]
names = [row['ticker'].replace('.BK', '') for _, row in df_sorted.iterrows()]

bars = ax.barh(names, df_sorted['esg_total'], color=bar_colors, alpha=0.85)
ax.axvline(x=70, color='gray', linestyle='--', alpha=0.5)
ax.set_xlim(0, 100)
ax.set_xlabel('ESG Score')
ax.set_title('ESG Ranking SET50 Top 10', fontsize=14, fontweight='bold')

for bar, score, rec in zip(bars, df_sorted['esg_total'], df_sorted['recommendation']):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            f"{score}  {rec}", va='center', fontsize=10, fontweight='bold')

plt.tight_layout()
plt.savefig('esg_ranking.png', dpi=150, bbox_inches='tight')
plt.show()
print("บันทึก esg_ranking.png แล้ว")


# Chart 2 — Investment Decision Matrix
fig, ax = plt.subplots(figsize=(12, 8))
val_map = {'Undervalued': 1, 'Fairly Valued': 2, 'Overvalued': 3}

offsets = {
    'CPALL': (-40, 8),  'SCB': (8, -12),
    'AOT':   (-32, 8),  'GULF': (-38, -12),
    'DELTA': (8, 8),    'KBANK': (8, 4),
    'SCC':   (8, 4),    'ADVANC': (8, 4),
    'BDMS':  (8, 4),    'PTT': (8, -12),
}

for _, row in df_dcf.iterrows():
    name = row['ticker'].replace('.BK', '')
    x = val_map.get(row['valuation'], 2)
    y = row['esg_total']
    color = colors_map.get(row['recommendation_esg'], '#3498db')
    xy_offset = offsets.get(name, (8, 4))

    ax.scatter(x, y, s=300, color=color, alpha=0.85, zorder=5)
    ax.annotate(name, (x, y),
                textcoords="offset points",
                xytext=xy_offset,
                fontsize=10, fontweight='bold')

ax.axhline(y=70, color='gray', linestyle='--', alpha=0.4)
ax.set_ylim(50, 90)
ax.set_xlim(0.5, 3.5)
ax.set_xticks([1, 2, 3])
ax.set_xticklabels(['Undervalued', 'Fairly Valued', 'Overvalued'], fontsize=11)
ax.set_ylabel('ESG Score', fontsize=12)
ax.set_title('Investment Decision Matrix\nESG Score vs DCF Valuation',
             fontsize=14, fontweight='bold')

ax.axvspan(0.5, 1.5, alpha=0.05, color='green')
ax.axvspan(1.5, 2.5, alpha=0.05, color='yellow')
ax.axvspan(2.5, 3.5, alpha=0.05, color='red')
ax.text(1, 88, 'Sweet Spot', ha='center', fontsize=9, color='green', style='italic')
ax.text(3, 88, 'Avoid Zone', ha='center', fontsize=9, color='red', style='italic')

patches = [
    mpatches.Patch(color='#2ecc71', label='BUY'),
    mpatches.Patch(color='#f39c12', label='HOLD'),
    mpatches.Patch(color='#e74c3c', label='AVOID')
]
ax.legend(handles=patches, loc='lower right')
plt.tight_layout()
plt.savefig('investment_matrix.png', dpi=150, bbox_inches='tight')
plt.show()
print("บันทึก investment_matrix.png แล้ว")


# ============================================================
# STEP 7 — Export Excel
# ============================================================

df_final = df_sorted.merge(
    df_dcf[['ticker', 'valuation', 'upside', 'dcf_summary']],
    on='ticker', how='left'
)
df_final.to_excel('esg_dcf_final.xlsx', index=False)
print("บันทึก esg_dcf_final.xlsx แล้ว")

print("\nสรุป Investment Decision Matrix:")
print("-" * 60)
for _, row in df_final.iterrows():
    name = row['ticker'].replace('.BK', '')
    print(f"{name:8} ESG:{row['esg_total']} {str(row.get('valuation','')):15} {str(row.get('upside','')):12} {row['recommendation']}")
