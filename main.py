from flask import Flask, request, jsonify, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_PATH = "製品マスタ.xlsx"  # 同じフォルダ内に置くこと

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    user_input = data.get("user_input", "")

    try:
        item_name = user_input.split("は")[0].strip()
        stock_str = user_input.split("は")[1].replace("個です", "").strip()
        stock = int(stock_str)
    except Exception:
        return jsonify({"error": "入力形式が正しくありません。例：トナーBは15個です"}), 400

    try:
        df = pd.read_excel(EXCEL_PATH)
        product_row = df[df["製品名"] == item_name].iloc[0]
        reorder_point = int(product_row["発注点"])
        reorder_qty = int(product_row["発注数"])
    except IndexError:
        return jsonify({"error": f"{item_name} は製品マスタに存在しません"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if stock <= reorder_point:
        status = "✅ 発注が必要です"
        order = reorder_qty
    else:
        status = "在庫は十分です"
        order = 0

    # Excel出力
    result_df = pd.DataFrame([{
        "製品名": item_name,
        "現在庫数": stock,
        "発注点": reorder_point,
        "発注要否": "要" if stock <= reorder_point else "不要",
        "発注数": order
    }])
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"発注リスト_{now}.xlsx"
    filepath = f"/tmp/{filename}"
    result_df.to_excel(filepath, index=False)

    return jsonify({
        "メッセージ": f"{item_name} の現在庫数は {stock} 個です。\n→ {status}（発注数：{order}個）",
        "ダウンロードURL": f"https://YOUR_RENDER_URL/download/{filename}"
    })

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    filepath = f"/tmp/{filename}"
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        return "ファイルが見つかりません", 404

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
