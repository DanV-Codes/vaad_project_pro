import argparse
import json
import os
from pathlib import Path
from excel_updater import LedgerDashboardReader, PAYING_APTS, MONTHLY_FEE, TOTAL_APTS, PETTY_AMOUNT

def generate_html(month_str, bank_balance, ledger_path, generate_png=True):
    print(f"Reading ledger: {ledger_path}")
    reader = LedgerDashboardReader(ledger_path)
    
    mm = month_str.split('/')[0]
    year = month_str.split('/')[1] if '/' in month_str else "2026"
    
    # 1. Income (הכנסות)
    target_income = PAYING_APTS * MONTHLY_FEE
    monthly_income = sum(
        apts.get(mm, 0) or 0 for apts in reader.payment_grid.values()
    )
    income_pct = (monthly_income / target_income) * 100 if target_income else 0
    
    # 2. Expenses (הוצאות)
    expense_items = []
    fixed_sum = 0
    for row in reader.fixed_rows:
        val = row["monthly"].get(mm)
        if val:
            expense_items.append({"name": row["category"], "amount": val})
            fixed_sum += val
            
    kelali_val = reader.kelali["monthly"].get(mm) or 0
    if kelali_val:
        expense_items.append({"name": reader.kelali["category"] or "כללי", "amount": kelali_val})
        
    total_expenses = sum(item["amount"] for item in expense_items)
    expense_comment = reader.kelali_comments.get(mm) or ""
    
    # 3. Expense Distribution (פילוג הוצאות)
    expense_items.sort(key=lambda x: x["amount"], reverse=True)
    
    # Take up to 5 items, group rest in "שאר הוצאות"
    top_expenses = expense_items[:5]
    other_expenses_amount = sum(item["amount"] for item in expense_items[5:])
    
    distribution = top_expenses.copy()
    if other_expenses_amount > 0 or not top_expenses:
        distribution.append({"name": "שאר ההוצאות", "amount": other_expenses_amount})
        
    # 4. Collection Percentages (אחוזי גביה מתחילת השנה)
    months_heb = {
        "01": "ינואר", "02": "פברואר", "03": "מרץ", "04": "אפריל",
        "05": "מאי", "06": "יוני", "07": "יולי", "08": "אוגוסט",
        "09": "ספטמבר", "10": "אוקטובר", "11": "נובמבר", "12": "דצמבר"
    }
    
    collection_history = []
    for m in range(1, int(mm) + 1):
        m_str = f"{m:02d}"
        m_income = sum(apts.get(m_str, 0) or 0 for apts in reader.payment_grid.values())
        pct = (m_income / target_income) * 100 if target_income else 0
        collection_history.append({"name": months_heb.get(m_str, m_str), "pct": pct})
    
    petty_pct = reader.petty_collection_pct
    collection_history.append({"name": "קופה קטנה", "pct": petty_pct})
    
    def get_color(pct):
        if pct >= 98: return "#10B981" # Green
        if pct >= 80: return "#FCD34D" # Yellow
        return "#EF4444" # Red
        
    month_name = months_heb.get(mm, mm)
    
    html = f"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>סיכום {month_name} {year}</title>
    <link href="https://fonts.googleapis.com/css2?family=Heebo:wght@700&display=swap" rel="stylesheet">
    <style>
        * {{ font-weight: 700; box-sizing: border-box; }}
        body {{
            font-family: 'Heebo', sans-serif;
            background-color: #F8FAFC;
            color: #1E293B;
            margin: 0;
            padding: 40px;
            display: flex;
            justify-content: center;
        }}
        .dashboard {{
            width: 1200px;
            display: flex;
            flex-direction: column;
            gap: 24px;
        }}
        .header {{
            text-align: center;
            font-size: 36px;
            margin-bottom: 20px;
        }}
        .row {{
            display: flex;
            gap: 24px;
        }}
        .card {{
            background: white;
            border-radius: 16px;
            padding: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
            flex: 1;
            display: flex;
            flex-direction: column;
        }}
        .card-title {{
            font-size: 20px;
            color: #000;
            margin-bottom: 16px;
            text-align: right;
        }}
        .card-value {{
            font-size: 46px;
            color: #334155;
            text-align: center;
            margin: 10px 0;
            display: flex;
            justify-content: center;
            align-items: baseline;
            gap: 8px;
        }}
        .currency {{
            font-size: 32px;
        }}
        .sub-text {{
            font-size: 15px;
            color: #475569;
            text-align: center;
            line-height: 1.4;
        }}
        .split-expenses {{
            font-size: 14px;
            color: #475569;
            text-align: center;
            margin-top: 5px;
        }}
        
        /* Progress Bar */
        .progress-bg {{
            background: #E2E8F0;
            height: 12px;
            border-radius: 6px;
            width: 100%;
            margin-top: 15px;
            overflow: hidden;
        }}
        .progress-fill {{
            background: #64748B;
            height: 100%;
            border-radius: 6px;
        }}
        
        /* Bar Chart */
        .chart-container {{
            display: flex;
            align-items: flex-end;
            justify-content: space-around;
            height: 250px;
            margin-top: 30px;
            position: relative;
            padding-bottom: 0px; 
            margin-bottom: 30px; /* space for x-labels */
        }}
        .grid-lines {{
            position: absolute;
            top: 0; left: 0; right: 0; bottom: 0;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            z-index: 0;
        }}
        .grid-line {{
            border-bottom: 1px solid #E2E8F0;
            width: 100%;
            position: relative;
        }}
        /* Base zero line */
        .grid-line-0 {{
            border-bottom: 2px solid #CBD5E1;
        }}
        .bar-wrapper {{
            display: flex;
            flex-direction: column;
            align-items: center;
            z-index: 1;
            height: 100%;
            justify-content: flex-end;
            width: 50px;
            position: relative;
        }}
        .bar-label {{
            position: absolute;
            bottom: -35px;
            font-size: 16px;
            color: #1E293B;
            white-space: nowrap;
        }}
        .bar-value {{
            font-size: 16px;
            margin-bottom: 8px;
            color: #1E293B;
        }}
        .bar {{
            width: 50px;
            border-radius: 6px 6px 0 0;
        }}
        
        /* Treemap / Mosaic */
        .mosaic-container {{
            display: flex;
            height: 300px;
            gap: 4px;
            border-radius: 12px;
            overflow: hidden;
            margin-top: 10px;
        }}
        .mosaic-col {{
            display: flex;
            flex-direction: column;
            gap: 4px;
            height: 100%;
        }}
        .mosaic-box {{
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            color: #1E293B;
            padding: 10px;
            text-align: center;
            gap: 6px;
        }}
        .box-title {{ 
            font-size: 18px; 
            line-height: 1.1;
        }}
        .box-val {{ 
            font-size: 16px; 
            line-height: 1.1;
        }}
    </style>
</head>
<body>

<div class="dashboard">
    <div class="header">האגמית 7 סיכום {month_name} {year}</div>
    
    <div class="row">
        <div class="card">
            <div class="card-title">יתרה בבנק</div>
            <div class="card-value">{bank_balance:,.0f} <span class="currency">₪</span></div>
            <div class="sub-text" style="margin-top: 15px;">נכון לסוף {month_name} {year}</div>
        </div>
        
        <div class="card">
            <div class="card-title" style="text-align: center;">הוצאות</div>
            <div class="card-value">{total_expenses:,.0f} <span class="currency">₪</span></div>
            <div class="split-expenses">
                שוטף: {fixed_sum:,.0f} ₪ <br>
                כללי: {kelali_val:,.0f} ₪ <span style="font-size:12px;">{f"({expense_comment})" if expense_comment else ""}</span>
            </div>
        </div>
        
        <div class="card">
            <div class="card-title">הכנסות</div>
            <div class="card-value">{monthly_income:,.0f} <span class="currency">₪</span></div>
            <div class="progress-bg">
                <div class="progress-fill" style="width: {min(100, income_pct)}%;"></div>
            </div>
            <div class="sub-text" style="margin-top: 12px;">{income_pct:.1f}% מתוך יעד של {target_income:,.0f} ₪</div>
        </div>
    </div>
    
    <div class="row">
        <div class="card" style="flex: 1;">
            <div class="card-title" style="text-align: center;">אחוזי גביה מתחילת השנה<br><span style="font-size: 14px; color: #64748B;">גרף התקדמות</span></div>
            
            <div class="chart-container">
                <div class="grid-lines">
                    <div class="grid-line" style="flex: 1;"></div>
                    <div class="grid-line" style="flex: 1;"></div>
                    <div class="grid-line" style="flex: 1;"></div>
                    <div class="grid-line" style="flex: 1;"></div>
                    <div class="grid-line grid-line-0" style="height: 0;"></div>
                </div>
                """
                
    for item in collection_history:
        pct_val = item["pct"]
        bar_h = min(100, max(0, pct_val))
        color = get_color(pct_val)
        html += f"""
                <div class="bar-wrapper">
                    <div class="bar-value">{pct_val:.1f}%</div>
                    <div class="bar" style="height: {bar_h}%; background-color: {color};"></div>
                    <div class="bar-label">{item["name"]}</div>
                </div>"""
                
    html += """
            </div>
        </div>
        
        <div class="card" style="flex: 1.2;">
            <div class="card-title" style="text-align: center;">פילוג הוצאות מרכזיות</div>
            <div class="mosaic-container">
"""
    # Mosaic logic for up to 6 items
    colors = ["#818CF8", "#A5B4FC", "#93C5FD", "#BAE6FD", "#E0F2FE", "#F1F5F9"]
    total_dist = sum(x["amount"] for x in distribution)
    
    def render_box(item, bg_color):
        return f'<div class="mosaic-box" style="height: 100%; background-color: {bg_color};"><div class="box-title">{item["name"]}</div><div class="box-val">(₪ {item["amount"]:,.0f})</div></div>'

    if total_dist == 0:
        html += '<div style="width: 100%; display: flex; align-items: center; justify-content: center; background: #F1F5F9;">אין הוצאות בחודש זה</div>'
    else:
        # Flexible multi-column layout
        chunks = []
        # Decide how many columns based on number of items
        if len(distribution) == 1:
            chunks = [[0]]
        elif len(distribution) == 2:
            chunks = [[0], [1]]
        elif len(distribution) == 3:
            chunks = [[0], [1, 2]]
        elif len(distribution) == 4:
            chunks = [[0], [1, 2], [3]]
        elif len(distribution) == 5:
            chunks = [[0], [1, 2], [3, 4]]
        else:
            chunks = [[0], [1, 2], [3, 4], [5]]
            
        for c_idx, chunk in enumerate(chunks):
            # calculate column width based on sum of amounts
            col_sum = sum(distribution[i]["amount"] for i in chunk)
            col_width_pct = max(15, (col_sum / total_dist) * 100) # give min 15% width
            html += f'<div class="mosaic-col" style="width: {col_width_pct}%;">'
            for item_idx in chunk:
                item = distribution[item_idx]
                item_h_pct = (item["amount"] / col_sum) * 100 if col_sum > 0 else 100
                html += f'<div style="height: {item_h_pct}%;">{render_box(item, colors[item_idx % len(colors)])}</div>'
            html += '</div>'

    html += """
            </div>
        </div>
    </div>
</div>
</body>
</html>
"""
    output_name = f"summary_{mm}_{year}.html"
    abs_path = os.path.abspath(output_name)
    with open(output_name, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Generated {output_name} successfully!")

    if generate_png:
        print("Attempting to generate PNG with Playwright...")
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page(device_scale_factor=2)  # High res
                page.goto(f"file:///{abs_path}")
                # Wait a tiny bit for fonts to load
                page.wait_for_timeout(500)
                # Take screenshot of the dashboard element
                element = page.locator(".dashboard")
                png_name = f"summary_{mm}_{year}.png"
                element.screenshot(path=png_name)
                print(f"Generated {png_name} successfully!")
                browser.close()
        except ImportError:
            print("Playwright is not installed. Run: pip install playwright && playwright install")
        except Exception as e:
            print(f"Failed to generate PNG: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate HTML Summary Slide")
    parser.add_argument("--month", required=True, help="Target month MM/YYYY (e.g., 05/2026)")
    parser.add_argument("--balance", type=float, required=True, help="Bank balance amount")
    parser.add_argument("--ledger", default="c:/Users/danie/PyProjects/vaad_project_pro_2_0/האגמית7_כספים_2026.xlsx", help="Path to master ledger")
    parser.add_argument("--no-png", action="store_true", help="Skip PNG generation")
    args = parser.parse_args()
    
    generate_html(args.month, args.balance, args.ledger, generate_png=not args.no_png)