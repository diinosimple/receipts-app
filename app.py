import os, base64, pickle, io
from flask import Flask, request, render_template
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from openpyxl import load_workbook

app = Flask(__name__)

# ===========================
# 設定値
# ===========================
TOKEN_PICKLE_B64 = "gASV+wMAAAAAAACMGWdvb2dsZS5vYXV0aDIuY3JlZGVudGlhbHOUjAtDcmVkZW50aWFsc5STlCmBlH2UKIwFdG9rZW6UjP15YTI5LmEwQVVNV2dfSjVoUDBOeWhCcENYWUFzUVk4cmphV3pJLXhoejhGT3FGYmdreEtudDNJUzZaVlNiX0VBRFpQcno2TVMzdmVZX09rY3BmTWxfb3Z1SVBfSV9ndE40LWItTVFESUIxamNXekhhOXNrNVktRjlWZG5NdGFlWTVDTzZJczlmQWFxZlhqZGVaRUYwTUZOZnFBYWlRZkU2QXdBRlNBcVV6dDl5UlhPdnV2ek9DcUI1UFJ3azdPVmhiZlpXVHJsTTZiWWItQWFDZ1lLQVJjU0FROFNGUUhHWDJNaUVWR0pQcTRfTHRpcXYwMDc1ektQMHcwMjA2lIwGZXhwaXJ5lIwIZGF0ZXRpbWWUjAhkYXRldGltZZSTlEMKB+oBEAoeOAAAAJSFlFKUjBFfcXVvdGFfcHJvamVjdF9pZJROjA9fdHJ1c3RfYm91bmRhcnmUTowQX3VuaXZlcnNlX2RvbWFpbpSMDmdvb2dsZWFwaXMuY29tlIwZX3VzZV9ub25fYmxvY2tpbmdfcmVmcmVzaJSJjAdfc2NvcGVzlF2UjCpodHRwczovL3d3dy5nb29nbGVhcGlzLmNvbS9hdXRoL2RyaXZlLmZpbGWUYYwPX2RlZmF1bHRfc2NvcGVzlE6MDl9yZWZyZXNoX3Rva2VulIxnMS8vMGdLM1pMY0puY1N5WENnWUlBUkFBR0JBU053Ri1MOUlyRkNVeGRvNTd5bGlzOWQyQnZCWkFNVWpzT3hMRVFVN3FHaW9NUDVvWUpFZFpOYUFTMDRkQnBfVU1INUd4a28za2tEMJSMCV9pZF90b2tlbpROjA9fZ3JhbnRlZF9zY29wZXOUXZSMKmh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL2F1dGgvZHJpdmUuZmlsZZRhjApfdG9rZW5fdXJplIwjaHR0cHM6Ly9vYXV0aDIuZ29vZ2xlYXBpcy5jb20vdG9rZW6UjApfY2xpZW50X2lklIxINzM4NTc5MzM3NDUzLWk0YjNraGFwNmYxMHJpcm84ajhjN2ZmamZyaDRlM2MwLmFwcHMuZ29vZ2xldXNlcmNvbnRlbnQuY29tlIwOX2NsaWVudF9zZWNyZXSUjCNHT0NTUFgtY3J5Y19CVVpjOVRZajFjbUUzc2o2VXJmR3M2epSMC19yYXB0X3Rva2VulE6MFl9lbmFibGVfcmVhdXRoX3JlZnJlc2iUiYwIX2FjY291bnSUjACUjA9fY3JlZF9maWxlX3BhdGiUTnViLg=="
EXCEL_FILE_ID = "1rf3DTxGpTNM0VZxcBkMjV2AyhE0oDiJlgv-_V_G3pbk"      # Excel ファイルID
RECEIPTS_FOLDER_ID = "1UaC4E-5O408ozxKx_VlFoYWilFWTbf-f"           # Drive フォルダID
SCOPES = ['https://www.googleapis.com/auth/drive']

# ===========================
# Google Drive サービス作成
# ===========================
def get_drive_service():
    if TOKEN_PICKLE_B64:
        token_bytes = base64.b64decode(TOKEN_PICKLE_B64)
        creds = pickle.load(io.BytesIO(token_bytes))
        service = build('drive', 'v3', credentials=creds)
        return service
    else:
        raise Exception("Google API credentials are missing")

# ===========================
# Excelに追記
# ===========================
def update_excel(service, filename, pay_date, payee, amount):
    request_dl = service.files().get_media(fileId=EXCEL_FILE_ID, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request_dl)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)

    wb = load_workbook(fh)
    ws = wb.active
    ws.append([filename, pay_date, payee, amount])

    out_fh = io.BytesIO()
    wb.save(out_fh)
    out_fh.seek(0)

    file_metadata = {"name": "receipts.xlsx"}
    media = MediaIoBaseUpload(out_fh, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
    service.files().update(fileId=EXCEL_FILE_ID, media_body=media, supportsAllDrives=True).execute()

# ===========================
# ファイルアップロード
# ===========================
def upload_file_to_drive(service, file, filename):
    file_metadata = {"name": filename, "parents": [RECEIPTS_FOLDER_ID]}
    media = MediaIoBaseUpload(file, mimetype=file.mimetype)
    service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()

# ===========================
# ルート
# ===========================
@app.route("/", methods=["GET", "POST"])
def index():
    message = ""
    if request.method == "POST":
        if "receipt" not in request.files:
            message = "画像が送信されていません。"
            return render_template("index.html", message=message)

        file = request.files["receipt"]
        if file.filename == "":
            message = "ファイル名がありません。"
            return render_template("index.html", message=message)

        pay_date = request.form.get("pay_date")
        payee = request.form.get("payee")
        amount = request.form.get("amount")

        # ファイル名作成
        filename = f"{payee} {pay_date} {amount}.jpg"

        try:
            service = get_drive_service()
            upload_file_to_drive(service, file, filename)
            update_excel(service, filename, pay_date, payee, amount)
            message = f"画像 {filename} を受信しました。"
        except Exception as e:
            message = f"エラー: {e}"

    return render_template("index.html", message=message)

# ===========================
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
