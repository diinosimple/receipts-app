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
TOKEN_PICKLE_B64 = "gASV+wMAAAAAAACMGWdvb2dsZS5vYXV0aDIuY3JlZGVudGlhbHOUjAtDcmVkZW50aWFsc5STlCmBlH2UKIwFdG9rZW6UjP15YTI5LmEwQVVNV2dfSjVoUDBOeWhCcENYWUFzUVk4cmphV3pJLXhoejhGT3FGYmdreEtudDNJUzZaVlNiX0VBRFpQcno2TVMzdmVZX09rY3BmTWxfb3Z1SVBfSV9ndE40LWItTVFESUIxamNXekhhOXNrNVktRjlWZG5NdGFlWTVDTzZJczlmQWFxZlhqZGVaRUYwTUZOZnFBYWlRZkU2QXdBRlNBcVV6dDl5UlhPdnV2ek9DcUI1UFJ3azdPVmhiZlpXVHJsTTZiWWItQWFDZ1lLQVJjU0FROFNGUUhHWDJNaUVWR0pQcTRfTHRpcXYwMDc1ektQMHcwMjA2lIwGZXhwaXJ5lIwIZGF0ZXRpbWWUjAhkYXRldGltZZSTlEMKB+oBEAoeOAAAAJSFlFKUjBFfcXVvdGFfcHJvamVjdF9pZJROjA9fdHJ1c3RfYm91bmRhcnmUTowQX3VuaXZlcnNlX2RvbWFpbpSMDmdvb2dsZWFwaXMuY29tlIwZX3VzZV9ub25fYmxvY2tpbmdfcmVmcmVzaJSJjAdfc2NvcGVzlF2UjCpodHRwczovL3d3dy5nb29nbGVhcGlzLmNvbS9hdXRoL2RyaXZlLmZpbGWUYYwPX2RlZmF1bHRfc2NvcGVzlE6MDl9yZWZyZXNoX3Rva2VulIxnMS8vMGdLM1pMY0puY1N5WENnWUlBUkFBR0JBU053Ri1MOUlyRkNVeGRvNTd5bGlzOWQyQnZCWkFNVWpzT3hMRVFVN3FHaW9NUDVvWUpFZFpOYUFTMDRkQnBfVU1INUd4a28za2tEMJSMCV9pZF90b2tlbpROjA9fZ3JhbnRlZF9zY29wZXOUXZSMKmh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL2F1dGgvZHJpdmUuZmlsZZRhjApfdG9rZW5fdXJplIwjaHR0cHM6Ly9vYXV0aDIuZ29vZ2xlYXBpcy5jb20vdG9rZW6UjApfY2xpZW50X2lklIxINzM4NTc5MzM3NDUzLWk0YjNraGFwNmYxMHJpcm84ajhjN2ZmamZyaDRlM2MwLmFwcHMuZ29vZ2xldXNlcmNvbnRlbnQuY29tlIwOX2NsaWVudF9zZWNyZXSUjCNHT0NTUFgtY3J5Y19CVVpjOVRZajFjbUUzc2o2VXJmR3M2epSMC19yYXB0X3Rva2VulE6MFl9lbmFibGVfcmVhdXRoX3JlZnJlc2iUiYwIX2FjY291bnSUjACUjA9fY3JlZF9maWxlX3BhdGiUTnViLg=="  # token.pickle を base64 にしたもの
EXCEL_FILE_ID = "1rf3DTxGpTNM0VZxcBkMjV2AyhE0oDiJlgv-_V_G3pbk"      # Excel ファイルID
RECEIPTS_FOLDER_ID = "1UaC4E-5O408ozxKx_VlFoYWilFWTbf-f"  # Drive フォルダID

print("EXCEL_FILE_ID:", EXCEL_FILE_ID)  # デバッグ用

# token.pickle から Credentials を復元
token_bytes = base64.b64decode(TOKEN_PICKLE_B64)
creds = pickle.load(io.BytesIO(token_bytes))

# -----------------------------
# Google Drive サービス作成
# -----------------------------
d# Drive サービス作成
def get_drive_service():
    try:
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        print("Error creating Drive service:", e)
        raise Exception("Google API credentials are invalid or missing")

# -----------------------------
# Excel ファイル取得/更新
# -----------------------------
ddef update_excel(service, filename, date_str, payee, amount):
    # ファイルをダウンロード
    request_dl = service.files().get_media(fileId=EXCEL_FILE_ID, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request_dl)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)

    # openpyxl で読み込み
    wb = load_workbook(fh)
    ws = wb.active
    ws.append([date_str, payee, amount, filename])

    # 一時保存
    temp_filename = "/tmp/temp.xlsx"
    wb.save(temp_filename)

    # Drive にアップロード（上書き）
    media = MediaFileUpload(temp_filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    service.files().update(fileId=EXCEL_FILE_ID, media_body=media, supportsAllDrives=True).execute()


# -----------------------------
# ルート
# -----------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        payee = request.form.get("payee")
        date_str = request.form.get("date")
        amount = request.form.get("amount")

        if not file:
            return "画像が送信されていません。", 400

        # 日付が空なら今日
        if not date_str:
            date_str = datetime.date.today().strftime("%Y-%m-%d")

        # ファイル名を作成
        filename = f"{payee} {date_str} {amount}.jpg"

        # Drive にアップロード
        service = get_drive_service()
        temp_path = f"/tmp/{filename}"
        file.save(temp_path)
        media = MediaFileUpload(temp_path, mimetype="image/jpeg")
        service.files().create(
            body={"name": filename, "parents": [RECEIPTS_FOLDER_ID]},
            media_body=media,
            supportsAllDrives=True
        ).execute()

        # Excel に追記
        update_excel(service, filename, date_str, payee, amount)

        return "画像を受信しました。"

    return render_template("index.html")
    
# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
