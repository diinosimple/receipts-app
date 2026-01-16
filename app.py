import os
import json
from datetime import datetime
from flask import Flask, render_template, request
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from openpyxl import Workbook, load_workbook
from io import BytesIO
import re

app = Flask(__name__)

# =========================
# Google Drive 設定
# =========================
SCOPES = ["https://www.googleapis.com/auth/drive"]
RECEIPTS_FOLDER_ID = os.environ["RECEIPTS_FOLDER_ID"]
EXCEL_FILE_ID = os.environ["EXCEL_FILE_ID"]

def get_drive_service():
    info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)

# =========================
# ユーティリティ
# =========================
def sanitize_filename(text):
    return re.sub(r'[\\/:*?"<>|]', "_", text)

def format_yen(amount):
    try:
        return f"¥{int(amount):,}"
    except:
        return "¥0"

# =========================
# Excel 追記
# =========================
def update_excel(service, row):
    try:
        request = service.files().get_media(fileId=EXCEL_FILE_ID)
        fh = BytesIO()
        request.execute(fh)
        fh.seek(0)
        wb = load_workbook(fh)
        ws = wb.active
    except Exception:
        wb = Workbook()
        ws = wb.active
        ws.append(["支払日", "支払い先", "金額", "内容"])

    ws.append(row)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    media = MediaIoBaseUpload(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )

    service.files().update(
        fileId=EXCEL_FILE_ID,
        media_body=media
    ).execute()

# =========================
# ルーティング
# =========================
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        image = request.files["image"]
        pay_date = request.form["pay_date"]
        vendor = request.form["vendor"]
        amount = request.form["amount"]
        description = request.form["description"]

        service = get_drive_service()

        yen = format_yen(amount)
        safe_vendor = sanitize_filename(vendor)
        filename = f"{safe_vendor} {pay_date} {yen}.jpg"

        media = MediaIoBaseUpload(
            image.stream,
            mimetype=image.mimetype,
            resumable=False
        )

        service.files().create(
            body={
                "name": filename,
                "parents": [RECEIPTS_FOLDER_ID]
            },
            media_body=media,
            fields="id"
        ).execute()

        update_excel(
            service,
            [pay_date, vendor, yen, description]
        )

        return "アップロード完了"

    return render_template("index.html")

# =========================
# Railway 用（port指定不要）
# =========================
if __name__ == "__main__":
    app.run()
