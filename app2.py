import os
import io
import base64
import pickle
from datetime import datetime

from flask import Flask, request, render_template, redirect
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.credentials import Credentials
from openpyxl import load_workbook

# =========================
# 環境変数
# =========================
EXCEL_FILE_ID = os.environ["EXCEL_FILE_ID"]        # Google Sheets ID
RECEIPTS_FOLDER_ID = os.environ["RECEIPTS_FOLDER_ID"]  # Drive フォルダID
TOKEN_PICKLE_B64 = os.environ["TOKEN_PICKLE_B64"]

SCOPES = ["https://www.googleapis.com/auth/drive"]

app = Flask(__name__)

# =========================
# Google Drive Service
# =========================
def get_drive_service():
    creds = pickle.loads(base64.b64decode(TOKEN_PICKLE_B64))
    return build("drive", "v3", credentials=creds)

# =========================
# ルート
# =========================
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        # ---------
        # 防御コード（Bad Request 対策）
        # ---------
        if "image" not in request.files:
            return "画像が送信されていません", 400

        image = request.files["image"]
        if image.filename == "":
            return "画像ファイルが空です", 400

        pay_date = request.form.get("pay_date")
        vendor = request.form.get("vendor")
        amount = request.form.get("amount")

        if not all([pay_date, vendor, amount]):
            return "フォーム項目が不足しています", 400

        drive = get_drive_service()


        # ==========
        # 画像アップロード
        # ==========
        filename = f"{vendor} {pay_date} ¥{amount}.jpg"

        media = MediaIoBaseUpload(
            image.stream,
            mimetype="image/jpeg",
            resumable=False
        )

        drive.files().create(
            body={
                "name": filename,
                "parents": [RECEIPTS_FOLDER_ID]
            },
            media_body=media,
            fields="id"
        ).execute()

        # ==========
        # Sheets を Excel として取得
        # ==========
        request_dl = drive.files().export(
            fileId=EXCEL_FILE_ID,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        excel_bytes = io.BytesIO(request_dl.execute())

        # ==========
        # Excel 追記
        # ==========
        wb = load_workbook(excel_bytes)
        ws = wb.active

        ws.append([
            pay_date,
            vendor,
            f"¥{amount}",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ])

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        # ==========
        # Drive に上書き保存
        # ==========
        media = MediaIoBaseUpload(
            out,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False
        )

        drive.files().update(
            fileId=EXCEL_FILE_ID,
            media_body=media
        ).execute()

        return redirect("/")

    return render_template("index.html")

# =========================
# Railway 用
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
