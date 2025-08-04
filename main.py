from flask import Flask, request, jsonify, send_file
import pandas as pd
from datetime import datetime
import os
import re

app = Flask(__name__)
EXCEL_PATH = "製品マスタ.xlsx"
EXPORT_DIR = "/tmp"

@app.route("/", methods=["GET"])
def index():
    return "✅ Dify在庫発注Webhookは稼働中です。"

@app.route("/webhook", methods=["POST"])
def webhook():
    # Content-Type 確認
    if not request.is_json:
        return jsonify({"response": "⚠️ JSON形式で送信してください（Content-Type: application/json）"}), 400

    data = request.get_json()
    if not data or "user_input" not in data:
        return jsonify({"response": "⚠️ user_input が見つかりません。例：トナーBは15個です"}), 400

    user_input = data["user_input"].strip()

    # 入力形式確認
    try:
        if "は" not in user_input or "個です" not in user_input:
            raise ValueError("形式不正")

        item_name, stock_part = user_input.split("は", 1)
        item_name = item_name.strip()
        stock_str = stock_part.replace("個です", "").strip()
        stock = int(stock_str)

    except Exception:
        return jsonify({"response": "⚠️ 入力形式が正しくありません。例：トナーBは15個です"}), 400

    try:
        df = pd.read_excel(EXCEL_PATH)
        if "製品名" not in df.columns or "発注点" not in df.columns or "発注数" not in df.columns:
            raise ValueError("マスタのカラムが不正")

        product_row = df[df["製品名"] == item_name]
        if product_row.empty:
            return jsonify({"response": f"⚠️ {item_name} は製品マスタに存在しません"}), 404

        reorder_point = int(product_row.iloc[0]["発注点"])
        reorder_qty = int(product_row.iloc[0]["発注数"])

    except Exception as e:
        return jsonify({"response": f"⚠️ マスタ読み込みエラー：{str(e)}"}), 500

    need_order = stock <= reorder_point
    order_status = "✅ 発注が必要です" if need_order else "在庫は十分です"
    order_qty = reorder_qty if need_order else 0

    result_df = pd.DataFrame([{
        "製品名": item_name,
        "現在庫数": stock,
        "発注点": reorder_point,
        "発注要否": "要" if need_order else "不要",
        "発注数": order_qty
    }])

    # ファイル名を安全に
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = re.sub(r"[^\w\-_.]", "_", f"発注リスト_{now}.xlsx")
    filepath = os.path.join(EXPORT_DIR, safe_name)

    try:
        result_df.to_excel(filepath, index=False)
    except Exception as e:
        return jsonify({"response": f"⚠️ Excel出力エラー：{str(e)}"}), 500

    download_url = f"https://dify-stock-order.onrender.com/download/{safe_name}"
    response_msg = (
        f"{item_name} の現在庫数は {stock} 個です。\n"
        f"→ {order_status}（発注数：{order_qty}個）\n\n"
        f"📎 ダウンロードURL：{download_url}"
    )
    return jsonify({"response": response_msg})

@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    safe_filename = re.sub(r"[^\w\-_.]", "_", filename)
    filepath = os.path.join(EXPORT_DIR, safe_filename)

    if not os.path.exists(filepath):
        return "⚠️ 指定されたファイルは存在しません", 404

    try:
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return f"⚠️ ファイル送信エラー：{str(e)}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
