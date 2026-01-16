import os
import base64
import pickle
import io
from datetime import datetime

from flask import Flask, request, render_template
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from openpyxl import load_workbook, Workbook

app = Flask(__name__)

# -----------------------------
# 環境変数（Railway 用）
# -----------------------------
COPES = ["https://www.googleapis.com/auth/drive"]
TOKEN_PICKLE_B64 = os.environ.get("TOKEN_PICKLE_B64")  # token.pickle を base64 にしたもの
EXCEL_FILE_ID = "1rf3DTxGpTNM0VZxcBkMjV2AyhE0oDiJlgv-_V_G3pbk"      # Excel ファイルID
RECEIPTS_FOLDER_ID = "1UaC4E-5O408ozxKx_VlFoYWilFWTbf-f"  # Drive フォルダID

print("EXCEL_FILE_ID:", EXCEL_FILE_ID)  # デバッグ用



# === OAuth 認証の取得 ===
def get_drive_service():
    """Google Drive API サービスオブジェクトを返す"""
    try:
        creds = None
        # TOKEN_PICKLE_B64 から creds を復元
        if 'TOKEN_PICKLE_B64' in os.environ:
            token_bytes = base64.b64decode(os.environ['TOKEN_PICKLE_B64'])
            creds = pickle.load(io.BytesIO(token_bytes))
        if not creds or not creds.valid:
            raise Exception("Google API credentials are invalid or missing")
        service = build('drive', 'v3', credentials=creds)
        print("Drive service successfully created")
        return service
    except Exception as e:
        print("Error creating Drive service:", e)
        raise  # Flask に伝えるため再度例外を投げる


    
# -----------------------------
# Excel ファイル取得/更新
# -----------------------------
def update_excel(service, filename, pay_date, payee, amount):
    # Excel ファイルを Drive からダウンロード
    request_dl = service.files().get_media(fileId=EXCEL_FILE_ID)
    fh = io.BytesIO(request_dl.execute())
    try:
        wb = load_workbook(fh)
    except:
        wb = Workbook()
    ws = wb.active

    # 末尾に追加
    ws.append([pay_date, payee, amount, filename])

    # 再び Drive にアップロード
    fh_upload = io.BytesIO()
    wb.save(fh_upload)
    fh_upload.seek(0)

    media = MediaIoBaseUpload(fh_upload, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
    service.files().update(fileId=EXCEL_FILE_ID, media_body=media).execute()



# -----------------------------
# ルート
# -----------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "image" not in request.files:
            return "画像が送信されていません。"

        file = request.files["image"]
        if file.filename == "":
            return "ファイル名が空です。"

        pay_date = request.form.get("pay_date", datetime.today().strftime("%Y-%m-%d"))
        payee = request.form.get("payee", "Unknown")
        amount = request.form.get("amount", "¥0")

        # ファイル名整形
        safe_payee = payee.replace(" ", "_")
        safe_amount = amount.replace(" ", "")
        filename = f"{safe_payee}_{pay_date}_{safe_amount}.jpg"

        # Drive にアップロード
        drive_service = get_drive_service()
        media = MediaIoBaseUpload(file, mimetype="image/jpeg")
        file_metadata = {
            "name": filename,
            "parents": [RECEIPTS_FOLDER_ID]
        }
        drive_service.files().create(body=file_metadata, media_body=media).execute()

        # Excel に追記
        update_excel(drive_service, filename, pay_date, payee, amount)

        return "画像を受信して Drive + Excel に反映しました 👍"

    return render_template("index.html")

# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
