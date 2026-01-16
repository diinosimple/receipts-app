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
SCOPES = ["https://www.googleapis.com/auth/drive"]
TOKEN_PICKLE_B64 = "gASV+wMAAAAAAACMGWdvb2dsZS5vYXV0aDIuY3JlZGVudGlhbHOUjAtDcmVkZW50aWFsc5STlCmBlH2UKIwFdG9rZW6UjP15YTI5LmEwQVVNV2dfSjVoUDBOeWhCcENYWUFzUVk4cmphV3pJLXhoejhGT3FGYmdreEtudDNJUzZaVlNiX0VBRFpQcno2TVMzdmVZX09rY3BmTWxfb3Z1SVBfSV9ndE40LWItTVFESUIxamNXekhhOXNrNVktRjlWZG5NdGFlWTVDTzZJczlmQWFxZlhqZGVaRUYwTUZOZnFBYWlRZkU2QXdBRlNBcVV6dDl5UlhPdnV2ek9DcUI1UFJ3azdPVmhiZlpXVHJsTTZiWWItQWFDZ1lLQVJjU0FROFNGUUhHWDJNaUVWR0pQcTRfTHRpcXYwMDc1ektQMHcwMjA2lIwGZXhwaXJ5lIwIZGF0ZXRpbWWUjAhkYXRldGltZZSTlEMKB+oBEAoeOAAAAJSFlFKUjBFfcXVvdGFfcHJvamVjdF9pZJROjA9fdHJ1c3RfYm91bmRhcnmUTowQX3VuaXZlcnNlX2RvbWFpbpSMDmdvb2dsZWFwaXMuY29tlIwZX3VzZV9ub25fYmxvY2tpbmdfcmVmcmVzaJSJjAdfc2NvcGVzlF2UjCpodHRwczovL3d3dy5nb29nbGVhcGlzLmNvbS9hdXRoL2RyaXZlLmZpbGWUYYwPX2RlZmF1bHRfc2NvcGVzlE6MDl9yZWZyZXNoX3Rva2VulIxnMS8vMGdLM1pMY0puY1N5WENnWUlBUkFBR0JBU053Ri1MOUlyRkNVeGRvNTd5bGlzOWQyQnZCWkFNVWpzT3hMRVFVN3FHaW9NUDVvWUpFZFpOYUFTMDRkQnBfVU1INUd4a28za2tEMJSMCV9pZF90b2tlbpROjA9fZ3JhbnRlZF9zY29wZXOUXZSMKmh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL2F1dGgvZHJpdmUuZmlsZZRhjApfdG9rZW5fdXJplIwjaHR0cHM6Ly9vYXV0aDIuZ29vZ2xlYXBpcy5jb20vdG9rZW6UjApfY2xpZW50X2lklIxINzM4NTc5MzM3NDUzLWk0YjNraGFwNmYxMHJpcm84ajhjN2ZmamZyaDRlM2MwLmFwcHMuZ29vZ2xldXNlcmNvbnRlbnQuY29tlIwOX2NsaWVudF9zZWNyZXSUjCNHT0NTUFgtY3J5Y19CVVpjOVRZajFjbUUzc2o2VXJmR3M2epSMC19yYXB0X3Rva2VulE6MFl9lbmFibGVfcmVhdXRoX3JlZnJlc2iUiYwIX2FjY291bnSUjACUjA9fY3JlZF9maWxlX3BhdGiUTnViLg=="  # token.pickle を base64 にしたもの
EXCEL_FILE_ID = "1rf3DTxGpTNM0VZxcBkMjV2AyhE0oDiJlgv-_V_G3pbk"      # Excel ファイルID
RECEIPTS_FOLDER_ID = "1UaC4E-5O408ozxKx_VlFoYWilFWTbf-f"  # Drive フォルダID

print("EXCEL_FILE_ID:", EXCEL_FILE_ID)  # デバッグ用


# -----------------------------
# Google Drive サービス作成
# -----------------------------
def get_drive_service():
    try:
        token_bytes = base64.b64decode(TOKEN_PICKLE_B64)
        creds = pickle.load(io.BytesIO(token_bytes))
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        print("Error creating Drive service:", e)
        raise Exception("Google API credentials are invalid or missing")

    
# -----------------------------
# Excel ファイル取得/更新
# -----------------------------
ddef update_excel(service, filename, pay_date, payee, amount):
    # Excel を Drive から取得
    request_dl = service.files().get_media(fileId=EXCEL_FILE_ID, supportsAllDrives=True)
    fh = io.BytesIO(request_dl.execute())
    wb = load_workbook(fh)
    ws = wb.active

    # 最終行に追加
    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), filename, pay_date, payee, amount])

    # Excel を再アップロード
    fh_out = io.BytesIO()
    wb.save(fh_out)
    fh_out.seek(0)
    media = MediaIoBaseUpload(fh_out, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    service.files().update(fileId=EXCEL_FILE_ID, media_body=media, supportsAllDrives=True).execute()


# -----------------------------
# ルート
# -----------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == 'POST':
        if 'receipt' not in request.files:
            return "画像が送信されていません。", 400

        file = request.files['receipt']
        payee = request.form.get('payee', 'Unknown')
        pay_date = request.form.get('pay_date', datetime.now().strftime("%Y-%m-%d"))
        amount = request.form.get('amount', '¥0')

        filename = f"{payee} {pay_date} {amount}.jpg"

        service = get_drive_service()

        # Drive にアップロード
        file_stream = io.BytesIO(file.read())
        media = MediaIoBaseUpload(file_stream, mimetype="image/jpeg")
        service.files().create(
            body={'name': filename, 'parents': [RECEIPTS_FOLDER_ID]},
            media_body=media,
            supportsAllDrives=True
        ).execute()

        # Excel に追記
        update_excel(service, filename, pay_date, payee, amount)

        return f"画像を受信しました: {filename}"

    return render_template('index.html')
    
# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
