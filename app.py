# app.py
from flask import Flask, request, render_template_string
import os
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)

# アップロード先（テスト用）
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# テンプレート（簡易版）
HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>領収書アップロード</title>
</head>
<body>
    <h2>領収書アップロード（iPhone対応）</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="receipt_image" accept="image/*" capture="environment" required>
        <br><br>
        支払日: <input type="date" name="pay_date" required><br><br>
        支払い先: <input type="text" name="payee" required><br><br>
        金額: <input type="text" name="amount" required><br><br>
        <button type="submit">アップロード</button>
    </form>
    {% if message %}
    <p>{{ message }}</p>
    {% endif %}
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    message = ""
    if request.method == "POST":
        # ファイル受信チェック
        if 'receipt_image' not in request.files:
            message = "画像が送信されていません。"
            return render_template_string(HTML_FORM, message=message), 400

        file = request.files['receipt_image']
        if file.filename == "":
            message = "画像が選択されていません。"
            return render_template_string(HTML_FORM, message=message), 400

        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)  # テスト用にサーバに保存
        print(f"画像を受信しました: {filepath}")

        # フォームデータ
        pay_date = request.form.get("pay_date")
        payee = request.form.get("payee")
        amount = request.form.get("amount")
        print(f"支払日: {pay_date}, 支払い先: {payee}, 金額: {amount}")

        # ここで Google Drive へのアップロードや Excel 反映を行う
        # upload_to_drive(filepath)
        # update_excel(pay_date, payee, amount, filepath)

        message = f"画像 '{filename}' を受信しました。"

    return render_template_string(HTML_FORM, message=message)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
