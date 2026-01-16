import os
import io
import base64
import pickle
from datetime import datetime
from flask import Flask, request, render_template
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from openpyxl import load_workbook

app = Flask(__name__)

# ------------------ 設定 ------------------
# Base64 に変換した token.pickle
TOKEN_PICKLE_B64 = "gASV+wMAAAAAAACMGWdvb2dsZS5vYXV0aDIuY3JlZGVudGlhbHOUjAtDcmVkZW50aWFsc5STlCmBlH2UKIwFdG9rZW6UjP15YTI5LmEwQVVNV2dfSjVoUDBOeWhCcENYWUFzUVk4cmphV3pJLXhoejhGT3FGYmdreEtudDNJUzZaVlNiX0VBRFpQcno2TVMzdmVZX09rY3BmTWxfb3Z1SVBfSV9ndE40LWItTVFESUIxamNXekhhOXNrNVktRjlWZG5NdGFlWTVDTzZJczlmQWFxZlhqZGVaRUYwTUZOZnFBYWlRZkU2QXdBRlNBcVV6dDl5UlhPdnV2ek9DcUI1UFJ3azdPVmhiZlpXVHJsTTZiWWItQWFDZ1lLQVJjU0FROFNGUUhHWDJNaUVWR0pQcTRfTHRpcXYwMDc1ektQMHcwMjA2lIwGZXhwaXJ5lIwIZGF0ZXRpbWWUjAhkYXRldGltZZSTlEMKB+oBEAoeOAAAAJSFlFKUjBFfcXVvdGFfcHJvamVjdF9pZJROjA9fdHJ1c3RfYm91bmRhcnmUTowQX3VuaXZlcnNlX2RvbWFpbpSMDmdvb2dsZWFwaXMuY29tlIwZX3VzZV9ub25fYmxvY2tpbmdfcmVmcmVzaJSJjAdfc2NvcGVzlF2UjCpodHRwczovL3d3dy5nb29nbGVhcGlzLmNvbS9hdXRoL2RyaXZlLmZpbGWUYYwPX2RlZmF1bHRfc2NvcGVzlE6MDl9yZWZyZXNoX3Rva2VulIxnMS8vMGdLM1pMY0puY1N5WENnWUlBUkFBR0JBU053Ri1MOUlyRkNVeGRvNTd5bGlzOWQyQnZCWkFNVWpzT3hMRVFVN3FHaW9NUDVvWUpFZFpOYUFTMDRkQnBfVU1INUd4a28za2tEMJSMCV9pZF90b2tlbpROjA9fZ3JhbnRlZF9zY29wZXOUXZSMKmh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL2F1dGgvZHJpdmUuZmlsZZRhjApfdG9rZW5fdXJplIwjaHR0cHM6Ly9vYXV0aDIuZ29vZ2xlYXBpcy5jb20vdG9rZW6UjApfY2xpZW50X2lklIxINzM4NTc5MzM3NDUzLWk0YjNraGFwNmYxMHJpcm84ajhjN2ZmamZyaDRlM2MwLmFwcHMuZ29vZ2xldXNlcmNvbnRlbnQuY29tlIwOX2NsaWVudF9zZWNyZXSUjCNHT0NTUFgtY3J5Y19CVVpjOVRZajFjbUUzc2o2VXJmR3M2epSMC19yYXB0X3Rva2VulE6MFl9lbmFibGVfcmVhdXRoX3JlZnJlc2iUiYwIX2FjY291bnSUjACUjA9fY3JlZF9maWxlX3BhdGiUTnViLg=="  # 実際は長い文字列
EXCEL_FILE_ID = "1rf3DTxGpTNM0VZxcBkMjV2AyhE0oDiJlgv-_V_G3pbk"
RECEIPTS_FOLDER_ID = "1UaC4E-5O408ozxKx_VlFoYWilFWTbf-f"
SCOPES = ["https://www.googleapis.com/auth/drive"]

# ------------------ Google Drive 認証 ------------------
def get_drive_service():
    token_bytes = base64.b64decode(TOKEN_PICKLE_B64)
    creds = pickle.load(io.BytesIO(token_bytes))
    service = build("drive", "v3", credentials=creds)
    return service

# ------------------ Excel 更新 ------------------
def update_excel(service, filename, pay_date, payee, amount):
    # Excel ファイルをダウンロード
    request_dl = service.files().get_media(fileId=EXCEL_FILE_ID)
    fh = io.BytesIO(request_dl.execute())

    # Excel を読み込み
    wb = load_workbook(filename=fh)
    ws = wb.active

    # 新しい行に追記
    ws.append([pay_date, payee, amount, filename])

    # 更新した Excel をバイトに保存
    out_fh = io.BytesIO()
    wb.save(out_fh)
    out_fh.seek(0)

    # Google Drive に上書きアップロード
    service.files().update(
        fileId=EXCEL_FILE_ID,
        media_body=MediaIoBaseUpload(out_fh, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ).execute()

# ------------------ Flask ルート ------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "receipt" not in request.files:
            return "画像が送信されていません。", 400

        file = request.files["receipt"]
        if file.filename == "":
            return "ファイル名がありません。", 400

        filename = file.filename
        now = datetime.now().strftime("%Y-%m-%d")
        payee = "テスト支払先"   # 実際はOCR等で取得
        amount = "¥1,000"        # 実際はOCR等で取得

        # Drive にアップロード
        service = get_drive_service()
        file_metadata = {
            "name": filename,
            "parents": [RECEIPTS_FOLDER_ID]
        }
        media = MediaFileUpload(file.filename, mimetype=file.mimetype)
        service.files().create(body=file_metadata, media_body=media, fields="id").execute()

        # Excel に追記
        update_excel(service, filename, now, payee, amount)

        return f"画像 {filename} を受信しました。"

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
