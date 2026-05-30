import os
import base64
import json
import re
from datetime import datetime
from io import BytesIO

from flask import Flask, request, jsonify, render_template, send_file
from dotenv import load_dotenv
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

WORK_CATEGORIES = [
    "クロス張替え（全面）",
    "クロス張替え（部分）",
    "クッションフロア張替え",
    "フローリング補修・部分貼替え",
    "畳替え",
    "ハウスクリーニング（基本）",
    "エアコンクリーニング",
    "浴室・強化クリーニング",
    "照明器具交換・修繕",
    "給湯器・設備修繕",
    "鍵交換",
    "リペア補修（床・壁複数箇所）",
    "残置物・廃棄物処分",
    "消毒・防虫処理",
    "諸経費（パーキング・交通費等）",
    "その他（自由入力）",
]

DEFAULT_BURDEN = {
    "クロス張替え（全面）": "貸",
    "クロス張替え（部分）": "借",
    "クッションフロア張替え": "貸",
    "フローリング補修・部分貼替え": "両",
    "畳替え": "両",
    "ハウスクリーニング（基本）": "借",
    "エアコンクリーニング": "借",
    "浴室・強化クリーニング": "借",
    "照明器具交換・修繕": "貸",
    "給湯器・設備修繕": "貸",
    "鍵交換": "借",
    "リペア補修（床・壁複数箇所）": "借",
    "残置物・廃棄物処分": "借",
    "消毒・防虫処理": "借",
    "諸経費（パーキング・交通費等）": "貸",
    "その他（自由入力）": "借",
}


@app.route("/")
def index():
    return render_template("index.html", categories=WORK_CATEGORIES)


@app.route("/api/analyze", methods=["POST"])
def analyze():
    if "pdf" not in request.files:
        return jsonify({"error": "PDFファイルが見つかりません"}), 400

    pdf_file = request.files["pdf"]
    if pdf_file.filename == "":
        return jsonify({"error": "ファイルが選択されていません"}), 400

    try:
        pdf_bytes = pdf_file.read()
        pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

        categories_str = "\n".join(f"- {c}" for c in WORK_CATEGORIES)

        prompt = f"""この見積書PDFを解析して、以下の情報をJSON形式で返してください。

抽出する情報：
1. 物件情報：
   - property_name: 物件名（不明な場合は空文字）
   - estimate_number: 見積番号（不明な場合は空文字）
   - estimate_date: 見積日（YYYY-MM-DD形式、不明な場合は空文字）
   - company_name: 業者名（不明な場合は空文字）
   - staff_name: 担当者名（不明な場合は空文字）

2. 工事明細（itemsリスト）：
   - category: 以下の工事分類から最も近いものを1つ選ぶ
{categories_str}
   - description: 仕様・備考（元の記載をそのまま）
   - quantity: 数量（数値）
   - unit: 単位（㎡、式、ヶ所、台など）
   - unit_price: 単価（税抜、数値）

レスポンスは必ず以下のJSON形式のみで返してください（説明文不要）：
{{
  "property_name": "",
  "estimate_number": "",
  "estimate_date": "",
  "company_name": "",
  "staff_name": "",
  "items": [
    {{
      "category": "",
      "description": "",
      "quantity": 0,
      "unit": "",
      "unit_price": 0
    }}
  ]
}}"""

        message = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=4096,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": pdf_b64,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt,
                        },
                    ],
                }
            ],
        )

        response_text = message.content[0].text.strip()

        # JSON部分を抽出
        json_match = re.search(r"\{.*\}", response_text, re.DOTALL)
        if json_match:
            response_text = json_match.group()

        result = json.loads(response_text)

        # 負担区分のデフォルト値を付与
        for item in result.get("items", []):
            category = item.get("category", "その他（自由入力）")
            item["burden"] = DEFAULT_BURDEN.get(category, "借")

        return jsonify(result)

    except json.JSONDecodeError as e:
        return jsonify({"error": f"APIレスポンスの解析に失敗しました: {str(e)}"}), 500
    except anthropic.APIError as e:
        return jsonify({"error": f"Claude APIエラー: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"error": f"エラーが発生しました: {str(e)}"}), 500


@app.route("/api/export", methods=["POST"])
def export():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "データが見つかりません"}), 400

        wb = create_excel(data)

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        property_name = data.get("property_name", "物件名不明")
        date_str = datetime.now().strftime("%Y%m%d")
        filename = f"原状回復精算_{property_name}_{date_str}.xlsx"

        # output/フォルダにも保存
        output_dir = os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(output_dir, exist_ok=True)
        save_path = os.path.join(output_dir, filename)
        with open(save_path, "wb") as f:
            output.seek(0)
            f.write(output.read())
        output.seek(0)

        encoded_filename = filename.encode("utf-8").decode("latin-1", errors="replace")

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as e:
        return jsonify({"error": f"Excel出力エラー: {str(e)}"}), 500


def make_border():
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)


def apply_header_style(cell, bg_color="1F3864"):
    cell.font = Font(name="Yu Gothic UI", bold=True, color="FFFFFF", size=10)
    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = make_border()


def apply_data_style(cell, bg_color=None, bold=False, align="left", number_format=None):
    cell.font = Font(name="Yu Gothic UI", bold=bold, size=10)
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = make_border()
    if number_format:
        cell.number_format = number_format


def create_excel(data):
    wb = openpyxl.Workbook()

    property_name = data.get("property_name", "")
    estimate_number = data.get("estimate_number", "")
    estimate_date = data.get("estimate_date", "")
    company_name = data.get("company_name", "")
    staff_name = data.get("staff_name", "")
    tenant_name = data.get("tenant_name", "")
    property_address = data.get("property_address", "")
    deposit = float(data.get("deposit", 0) or 0)
    landlord_rate = float(data.get("landlord_rate", 1.15) or 1.15)
    tenant_rate = float(data.get("tenant_rate", 1.20) or 1.20)
    tax_rate = float(data.get("tax_rate", 0.10) or 0.10)
    items = data.get("items", [])

    # ① 負担割合表シート
    ws1 = wb.active
    ws1.title = "負担割合表"
    _build_estimate_sheet(ws1, data, items, landlord_rate, tenant_rate, tax_rate,
                          property_name, estimate_number, estimate_date, company_name, staff_name,
                          tenant_name, property_address)

    # ② 貸主請求書
    ws2 = wb.create_sheet("貸主請求書")
    landlord_items = [i for i in items if i.get("burden") in ("貸", "両")]
    _build_landlord_sheet(ws2, landlord_items, data, landlord_rate, tax_rate)

    # ③ 借主清算書
    ws3 = wb.create_sheet("借主清算書")
    tenant_items = [i for i in items if i.get("burden") in ("借", "両")]
    _build_tenant_sheet(ws3, tenant_items, data, tenant_rate, tax_rate, deposit)

    return wb


COMPANY_INFO = "恵比寿不動産（株式会社ライフアドバンス）　賃貸管理事業部　〒150-0011 東京都渋谷区東3丁目25-11 TOKYU REIT恵比寿ビル4F　TEL: 03-6421-0544"


def _build_estimate_sheet(ws, data, items, landlord_rate, tenant_rate, tax_rate,
                           property_name, estimate_number, estimate_date, company_name, staff_name,
                           tenant_name="", property_address=""):
    # 12列構成: A=No B=名称 C=施工箇所 D=数量 E=単位 F=単価 G=金額
    #           H=借主割合 I=借主金額 J=貸主割合 K=貸主金額 L=備考
    col_w = {"A":5,"B":22,"C":14,"D":7,"E":5,"F":10,"G":12,"H":7,"I":12,"J":7,"K":12,"L":20}
    for col, w in col_w.items():
        ws.column_dimensions[col].width = w

    issue_date = datetime.now().strftime("%Y年%m月%d日")

    # ── Row1: 発行日 ──────────────────────────────
    ws["A1"] = f"発行日：{issue_date}"
    ws["A1"].font = Font(name="Yu Gothic UI", size=9, color="555555")
    ws.row_dimensions[1].height = 14

    # ── Row2: タイトル（中央）＋ 会社名（右） ────────────
    ws.merge_cells("C2:G2")
    c = ws["C2"]; c.value = "お見積書"
    c.font = Font(name="Yu Gothic UI", bold=True, size=20, color="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("H2:L2")
    c = ws["H2"]; c.value = "株式会社 ライフアドバンス"
    c.font = Font(name="Yu Gothic UI", bold=True, size=11, color="1F3864")
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 34

    # ── Row3: 借主名（左）＋ 会社住所（右） ──────────────
    ws.merge_cells("A3:E3")
    c = ws["A3"]; c.value = f"{tenant_name}　様" if tenant_name else "　様"
    c.font = Font(name="Yu Gothic UI", bold=True, size=14)
    c.alignment = Alignment(horizontal="left", vertical="bottom")
    ws.merge_cells("H3:L3")
    c = ws["H3"]; c.value = "所在地：東京都渋谷区東3-25-11 TOKYU REIT恵比寿ビル4階"
    c.font = Font(name="Yu Gothic UI", size=8, color="666666")
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[3].height = 24

    # ── Row4: お見積もり金額（貸主負担税込） ─────────────
    landlord_ex_preview = 0
    for item in items:
        qty = float(item.get("quantity", 0) or 0)
        price = float(item.get("unit_price", 0) or 0)
        vendor = int(qty * price)
        burden = item.get("burden", "借")
        if burden == "貸":
            landlord_ex_preview += int(vendor * landlord_rate)
        elif burden == "両":
            landlord_ex_preview += int((vendor / 2) * landlord_rate)
    landlord_tax_preview = int(landlord_ex_preview * (1 + tax_rate))

    ws.merge_cells("A4:B4")
    c = ws["A4"]; c.value = "お見積もり金額(税込)"
    c.font = Font(name="Yu Gothic UI", size=9, bold=True)
    c.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = make_border()
    ws.merge_cells("C4:F4")
    c = ws["C4"]; c.value = landlord_tax_preview
    c.number_format = '#,##0"円"'
    c.font = Font(name="Yu Gothic UI", bold=True, size=14, color="1F3864")
    c.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = make_border()
    ws.row_dimensions[4].height = 26

    # ── Row5-7: 工事情報 ──────────────────────────
    prop_rows = [("工事件名", "原状回復工事"), ("物件名", property_name), ("物件住所", property_address)]
    r = 5
    for label, value in prop_rows:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        c = ws.cell(row=r, column=1, value=label)
        c.font = Font(name="Yu Gothic UI", size=9, bold=True)
        c.fill = PatternFill(start_color="E8EEF8", end_color="E8EEF8", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = make_border()
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        c = ws.cell(row=r, column=3, value=value)
        c.font = Font(name="Yu Gothic UI", size=9)
        c.border = make_border()
        ws.row_dimensions[r].height = 16
        r += 1

    r += 1  # 空行
    ws.row_dimensions[r - 1].height = 6

    # ── テーブルヘッダー（2行） ───────────────────────
    # Row1: 単一セルは2行スパン、借主/貸主は2列スパン
    row_h1 = r
    for col, val in [(1,"No."),(2,"名称"),(3,"施工箇所"),(4,"数量"),(5,"単位"),(6,"単価"),(7,"金額"),(12,"備考")]:
        ws.merge_cells(start_row=row_h1, start_column=col, end_row=row_h1+1, end_column=col)
        c = ws.cell(row=row_h1, column=col, value=val)
        apply_header_style(c, "1F3864")
    ws.merge_cells(start_row=row_h1, start_column=8, end_row=row_h1, end_column=9)
    c = ws.cell(row=row_h1, column=8, value="借主様負担額")
    apply_header_style(c, "8B2500")
    ws.merge_cells(start_row=row_h1, start_column=10, end_row=row_h1, end_column=11)
    c = ws.cell(row=row_h1, column=10, value="貸主様負担額")
    apply_header_style(c, "1F3864")
    ws.row_dimensions[row_h1].height = 18

    # Row2: 割合/金額
    row_h2 = row_h1 + 1
    for col, val, bg in [(8,"割合","8B2500"),(9,"金額","8B2500"),(10,"割合","1F3864"),(11,"金額","1F3864")]:
        c = ws.cell(row=row_h2, column=col, value=val)
        apply_header_style(c, bg)
    ws.row_dimensions[row_h2].height = 16
    r = row_h2 + 1

    # ── データ行（20行固定） ──────────────────────────
    vendor_total = 0
    tenant_total_ex = 0
    landlord_total_ex = 0

    for i in range(1, 21):
        if i <= len(items):
            item = items[i - 1]
            qty = float(item.get("quantity", 0) or 0)
            price = float(item.get("unit_price", 0) or 0)
            vendor = int(qty * price)
            burden = item.get("burden", "借")
            if burden == "貸":
                t_ratio, l_ratio = "0%", "100%"
                t_amt = 0
                l_amt = int(vendor * landlord_rate)
            elif burden == "借":
                t_ratio, l_ratio = "100%", "0%"
                t_amt = int(vendor * tenant_rate)
                l_amt = 0
            elif burden == "両":
                t_ratio, l_ratio = "50%", "50%"
                t_amt = int((vendor / 2) * tenant_rate)
                l_amt = int((vendor / 2) * landlord_rate)
            else:
                t_ratio, l_ratio = "", ""
                t_amt, l_amt = 0, 0
            vendor_total += vendor
            tenant_total_ex += t_amt
            landlord_total_ex += l_amt
            bg = "FCE4D6" if burden == "借" else ("D9E1F2" if burden == "貸" else "E2EFDA")
            row_vals = [i, item.get("category",""), item.get("description",""),
                        qty, item.get("unit",""), price, vendor,
                        t_ratio, t_amt, l_ratio, l_amt, ""]
        else:
            bg = None
            row_vals = [i,"","","","","","","","","","",""]

        for col, val in enumerate(row_vals, 1):
            c = ws.cell(row=r, column=col, value=val if (i <= len(items) or col == 1) else None)
            c.font = Font(name="Yu Gothic UI", size=10)
            if bg:
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.border = make_border()
            if col in (6, 7, 9, 11) and i <= len(items):
                c.number_format = "#,##0"
            c.alignment = Alignment(
                horizontal="center" if col in (1, 4, 5, 8, 10) else ("right" if col in (6,7,9,11) else "left"),
                vertical="center"
            )
        ws.row_dimensions[r].height = 18
        r += 1

    # ── 小計・消費税・合計 ────────────────────────────
    tax_pct = int(tax_rate * 100)
    totals = [
        ("小計", vendor_total, tenant_total_ex, landlord_total_ex),
        (f"消費税　{tax_pct}%", int(vendor_total*tax_rate), int(tenant_total_ex*tax_rate), int(landlord_total_ex*tax_rate)),
        ("合計", int(vendor_total*(1+tax_rate)), int(tenant_total_ex*(1+tax_rate)), int(landlord_total_ex*(1+tax_rate))),
    ]
    for label, v_val, t_val, l_val in totals:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c = ws.cell(row=r, column=1, value=label)
        c.font = Font(name="Yu Gothic UI", bold=True, size=10)
        c.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = make_border()
        for col, val in [(7,v_val),(9,t_val),(11,l_val)]:
            c = ws.cell(row=r, column=col, value=val)
            c.number_format = '#,##0"円"'
            apply_data_style(c, bg_color="FFF2CC", bold=True, align="right")
        for col in [8, 10, 12]:
            c = ws.cell(row=r, column=col)
            c.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            c.border = make_border()
        ws.row_dimensions[r].height = 18
        r += 1

    # ── 弊社利益（内部確認用） ──────────────────────────
    profit = (landlord_total_ex + tenant_total_ex) - vendor_total
    profit_rate = (profit / vendor_total * 100) if vendor_total > 0 else 0
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = ws.cell(row=r, column=1, value=f"【内部】弊社利益（税抜）：¥{profit:,}　利益率：{profit_rate:.1f}%")
    apply_data_style(c, bg_color="FFE699", bold=True, align="center")
    ws.row_dimensions[r].height = 18
    r += 2

    # ── 備考 ──────────────────────────────────────
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = ws.cell(row=r, column=1, value="【備考】")
    c.font = Font(name="Yu Gothic UI", bold=True, size=10)
    c.border = make_border()
    ws.row_dimensions[r].height = 16
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=12)
    c = ws.cell(row=r, column=1, value="")
    c.border = make_border()
    ws.row_dimensions[r].height = 20

    ws.freeze_panes = ws.cell(row=row_h2 + 1, column=1)


def _build_landlord_sheet(ws, items, data, rate, tax_rate):
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 6
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 15

    ws.merge_cells("A1:H1")
    co_cell = ws["A1"]
    co_cell.value = COMPANY_INFO
    co_cell.font = Font(name="Yu Gothic UI", size=9, color="444444")
    co_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 16

    ws.merge_cells("A2:H2")
    title = ws["A2"]
    title.value = "貸主請求書"
    title.font = Font(name="Yu Gothic UI", bold=True, size=14, color="1F3864")
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 30

    row = 3
    info_labels = [("物件名", data.get("property_name", "")), ("見積番号", data.get("estimate_number", "")),
                   ("業者名", data.get("company_name", "")), ("担当者", data.get("staff_name", ""))]
    for label, value in info_labels:
        ws.cell(row=row, column=1, value=label).font = Font(name="Yu Gothic UI", bold=True, size=10)
        ws.cell(row=row, column=1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
        ws.cell(row=row, column=2, value=value).font = Font(name="Yu Gothic UI", size=10)
        row += 1

    row += 1

    headers = ["No", "工事項目", "仕様備考", "数量", "単位", "原価（税抜）", "請求額（税抜）", "請求額（税込）"]
    for col, h in enumerate(headers, 1):
        apply_header_style(ws.cell(row=row, column=col), "1F3864")
        ws.cell(row=row, column=col).value = h
    ws.row_dimensions[row].height = 22
    row += 1

    subtotal_ex = 0
    for i, item in enumerate(items, 1):
        quantity = float(item.get("quantity", 0) or 0)
        unit_price = float(item.get("unit_price", 0) or 0)
        vendor_amount = int(quantity * unit_price)

        burden = item.get("burden", "貸")
        if burden == "両":
            vendor_amount = vendor_amount // 2

        charge_ex = int(vendor_amount * rate)
        charge_tax = int(charge_ex * (1 + tax_rate))
        subtotal_ex += charge_ex

        row_data = [i, item.get("category", ""), item.get("description", ""),
                    quantity, item.get("unit", ""), vendor_amount, charge_ex, charge_tax]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row, column=col, value=val)
            apply_data_style(cell, bg_color="D9E1F2",
                             align="center" if col in (1, 4, 5) else ("right" if col >= 6 else "left"))
            if col in (6, 7, 8):
                cell.number_format = "#,##0"
        ws.row_dimensions[row].height = 18
        row += 1

    total_tax = int(subtotal_ex * (1 + tax_rate))

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value="合計")
    apply_data_style(cell, bg_color="FFF2CC", bold=True, align="center")
    for col, val in [(7, subtotal_ex), (8, total_tax)]:
        cell = ws.cell(row=row, column=col, value=val)
        apply_data_style(cell, bg_color="FFF2CC", bold=True, align="right", number_format="#,##0")


def _build_tenant_sheet(ws, items, data, rate, tax_rate, deposit):
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 6
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 15

    ws.merge_cells("A1:H1")
    co_cell = ws["A1"]
    co_cell.value = COMPANY_INFO
    co_cell.font = Font(name="Yu Gothic UI", size=9, color="444444")
    co_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 16

    ws.merge_cells("A2:H2")
    title = ws["A2"]
    title.value = "借主清算書"
    title.font = Font(name="Yu Gothic UI", bold=True, size=14, color="1F3864")
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 30

    row = 3
    info_labels = [("物件名", data.get("property_name", "")), ("見積番号", data.get("estimate_number", "")),
                   ("業者名", data.get("company_name", "")), ("担当者", data.get("staff_name", ""))]
    for label, value in info_labels:
        ws.cell(row=row, column=1, value=label).font = Font(name="Yu Gothic UI", bold=True, size=10)
        ws.cell(row=row, column=1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
        ws.cell(row=row, column=2, value=value).font = Font(name="Yu Gothic UI", size=10)
        row += 1

    row += 1

    headers = ["No", "工事項目", "仕様備考", "数量", "単位", "原価（税抜）", "請求額（税抜）", "請求額（税込）"]
    for col, h in enumerate(headers, 1):
        apply_header_style(ws.cell(row=row, column=col), "1F3864")
        ws.cell(row=row, column=col).value = h
    ws.row_dimensions[row].height = 22
    row += 1

    subtotal_ex = 0
    for i, item in enumerate(items, 1):
        quantity = float(item.get("quantity", 0) or 0)
        unit_price = float(item.get("unit_price", 0) or 0)
        vendor_amount = int(quantity * unit_price)

        burden = item.get("burden", "借")
        if burden == "両":
            vendor_amount = vendor_amount // 2

        charge_ex = int(vendor_amount * rate)
        charge_tax = int(charge_ex * (1 + tax_rate))
        subtotal_ex += charge_ex

        row_data = [i, item.get("category", ""), item.get("description", ""),
                    quantity, item.get("unit", ""), vendor_amount, charge_ex, charge_tax]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row, column=col, value=val)
            apply_data_style(cell, bg_color="FCE4D6",
                             align="center" if col in (1, 4, 5) else ("right" if col >= 6 else "left"))
            if col in (6, 7, 8):
                cell.number_format = "#,##0"
        ws.row_dimensions[row].height = 18
        row += 1

    total_tax = int(subtotal_ex * (1 + tax_rate))

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value="工事費合計（税込）")
    apply_data_style(cell, bg_color="FFF2CC", bold=True, align="center")
    ws.cell(row=row, column=7, value="").fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws.cell(row=row, column=7).border = make_border()
    cell = ws.cell(row=row, column=8, value=total_tax)
    apply_data_style(cell, bg_color="FFF2CC", bold=True, align="right", number_format="#,##0")
    row += 1

    # 敷金精算
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value="敷金（預かり金）")
    apply_data_style(cell, bg_color="FFF2CC", bold=True, align="center")
    ws.cell(row=row, column=7, value="").fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws.cell(row=row, column=7).border = make_border()
    cell = ws.cell(row=row, column=8, value=int(deposit))
    apply_data_style(cell, bg_color="FFF2CC", bold=True, align="right", number_format="#,##0")
    row += 1

    balance = int(deposit) - total_tax
    if balance >= 0:
        label = f"返金額：¥{balance:,}"
        bg = "E2EFDA"
    else:
        label = f"追加請求額：¥{abs(balance):,}"
        bg = "FCE4D6"

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    cell = ws.cell(row=row, column=1, value=label)
    apply_data_style(cell, bg_color=bg, bold=True, align="center")
    cell = ws.cell(row=row, column=8, value=abs(balance))
    apply_data_style(cell, bg_color=bg, bold=True, align="right", number_format="#,##0")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, use_reloader=False)
