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
    deposit = float(data.get("deposit", 0) or 0)
    landlord_rate = float(data.get("landlord_rate", 1.15) or 1.15)
    tenant_rate = float(data.get("tenant_rate", 1.20) or 1.20)
    tax_rate = float(data.get("tax_rate", 0.10) or 0.10)
    items = data.get("items", [])

    # ① 見積入力シート
    ws1 = wb.active
    ws1.title = "見積入力"
    _build_estimate_sheet(ws1, data, items, landlord_rate, tenant_rate, tax_rate,
                          property_name, estimate_number, estimate_date, company_name, staff_name)

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
                           property_name, estimate_number, estimate_date, company_name, staff_name):
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 6
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 13
    ws.column_dimensions["H"].width = 8
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["J"].width = 14

    # 発行元会社情報
    ws.merge_cells("A1:J1")
    co_cell = ws["A1"]
    co_cell.value = COMPANY_INFO
    co_cell.font = Font(name="Yu Gothic UI", size=9, color="444444")
    co_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 16

    # タイトル
    ws.merge_cells("A2:J2")
    title_cell = ws["A2"]
    title_cell.value = "原状回復精算 見積入力シート"
    title_cell.font = Font(name="Yu Gothic UI", bold=True, size=14, color="1F3864")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 30

    # 物件情報
    info = [
        ("物件名", property_name),
        ("見積番号", estimate_number),
        ("見積日", estimate_date),
        ("業者名", company_name),
        ("担当者", staff_name),
    ]
    row = 3
    for label, value in info:
        ws.cell(row=row, column=1, value=label).font = Font(name="Yu Gothic UI", bold=True, size=10)
        ws.cell(row=row, column=1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
        ws.cell(row=row, column=2, value=value).font = Font(name="Yu Gothic UI", size=10)
        ws.row_dimensions[row].height = 18
        row += 1

    row += 1

    # ヘッダー
    headers = ["No", "工事項目", "仕様備考", "数量", "単位", "単価（税抜）", "業者金額", "負担区分", "貸主請求額", "借主請求額"]
    for col, h in enumerate(headers, 1):
        apply_header_style(ws.cell(row=row, column=col), "1F3864")
        ws.cell(row=row, column=col).value = h
    ws.row_dimensions[row].height = 22
    header_row = row
    row += 1

    landlord_total = 0
    tenant_total = 0
    vendor_total = 0

    for i, item in enumerate(items, 1):
        category = item.get("category", "")
        description = item.get("description", "")
        quantity = float(item.get("quantity", 0) or 0)
        unit = item.get("unit", "")
        unit_price = float(item.get("unit_price", 0) or 0)
        burden = item.get("burden", "借")

        vendor_amount = int(quantity * unit_price)
        landlord_amount = 0
        tenant_amount = 0

        if burden == "貸":
            landlord_amount = int(vendor_amount * landlord_rate)
        elif burden == "借":
            tenant_amount = int(vendor_amount * tenant_rate)
        elif burden == "両":
            landlord_amount = int((vendor_amount / 2) * landlord_rate)
            tenant_amount = int((vendor_amount / 2) * tenant_rate)

        landlord_total += landlord_amount
        tenant_total += tenant_amount
        vendor_total += vendor_amount

        if burden == "貸":
            bg = "D9E1F2"
        elif burden == "借":
            bg = "FCE4D6"
        else:
            bg = "E2EFDA"

        row_data = [i, category, description, quantity, unit, unit_price, vendor_amount, burden, landlord_amount, tenant_amount]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row, column=col, value=val)
            apply_data_style(cell, bg_color=bg, align="center" if col in (1, 4, 5, 8) else ("right" if col >= 6 else "left"))
            if col in (6, 7, 9, 10):
                cell.number_format = "#,##0"

        ws.row_dimensions[row].height = 18
        row += 1

    # 合計行
    total_row = row
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=6)
    cell = ws.cell(row=total_row, column=1, value="合計（税抜）")
    apply_data_style(cell, bg_color="FFF2CC", bold=True, align="center")

    for col, val in [(7, vendor_total), (9, landlord_total), (10, tenant_total)]:
        cell = ws.cell(row=total_row, column=col, value=val)
        apply_data_style(cell, bg_color="FFF2CC", bold=True, align="right", number_format="#,##0")

    ws.cell(row=total_row, column=8, value="").fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws.cell(row=total_row, column=8).border = make_border()

    row += 1

    # 税込合計
    landlord_tax = int(landlord_total * (1 + tax_rate))
    tenant_tax = int(tenant_total * (1 + tax_rate))
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value=f"合計（税込 {int(tax_rate*100)}%）")
    apply_data_style(cell, bg_color="FFF2CC", bold=True, align="center")
    ws.cell(row=row, column=8, value="").fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws.cell(row=row, column=8).border = make_border()
    for col, val in [(7, int(vendor_total * (1 + tax_rate))), (9, landlord_tax), (10, tenant_tax)]:
        cell = ws.cell(row=row, column=col, value=val)
        apply_data_style(cell, bg_color="FFF2CC", bold=True, align="right", number_format="#,##0")

    row += 1

    # 利益
    profit = (landlord_total + tenant_total) - vendor_total
    profit_rate = (profit / vendor_total * 100) if vendor_total > 0 else 0
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws.cell(row=row, column=1, value=f"弊社利益（税抜）：¥{profit:,}　利益率：{profit_rate:.1f}%")
    apply_data_style(cell, bg_color="FFE699", bold=True, align="center")
    ws.cell(row=row, column=9, value="").fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    ws.cell(row=row, column=9).border = make_border()
    ws.cell(row=row, column=10, value="").fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    ws.cell(row=row, column=10).border = make_border()

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)


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
