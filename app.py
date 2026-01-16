from flask import Flask, request, render_template

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # デバッグ用ログ（Railway / ローカルどちらでも重要）
        print("==== DEBUG START ====")
        print("request.content_type:", request.content_type)
        print("request.files:", request.files)
        print("request.form:", request.form)
        print("==== DEBUG END ====")

        if "image" not in request.files:
            return "画像が送信されていません。"

        file = request.files["image"]

        if file.filename == "":
            return "ファイル名が空です。"

        # ここでは保存せず、受信確認のみ
        return "画像を受信しました 👍"

    return render_template("index.html")


if __name__ == "__main__":
    # ローカルテスト用
    app.run(host="0.0.0.0", port=5001, debug=True)
