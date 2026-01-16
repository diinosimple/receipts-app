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
TOKEN_PICKLE_B64 = "gASV8QMAAAAAAACMGWdvb2dsZS5vYXV0aDIuY3JlZGVudGlhbHOUjAtDcmVkZW50aWFsc5STlCmBlH2UKIwFdG9rZW6UjP15YTI5LmEwQVVNV2dfTGFYeTQ5dC1pb082Y1JrRVMxcGJhRUkxUE5HVnlkMlBNZnA2MGQtMUtOdWdIV0VwejFiS0NYU0JvVFY3aEtWT19NektTTUdLbV9lQ2dPd1J5UG9IT2RiTk5WQV9lTmF5cjNUMlhUaDd1Nmx0Z0FNTkNPcDYyV2hOdlA4bHNzbnlPbEdrc0RKNFZCRWowZzE4UVBVY0pCTUNFZDg0UTZvWVFKZVZzaTlhb2J6VUM0bnAtaFQyZ3RjVVRvc3pNWG10c2FDZ1lLQVZZU0FROFNGUUhHWDJNaUctMWZLQjlreXk5cDlOTk1CdW4wQVEwMjA2lIwGZXhwaXJ5lIwIZGF0ZXRpbWWUjAhkYXRldGltZZSTlEMKB+oBEA4TOAAAAJSFlFKUjBFfcXVvdGFfcHJvamVjdF9pZJROjA9fdHJ1c3RfYm91bmRhcnmUTowQX3VuaXZlcnNlX2RvbWFpbpSMDmdvb2dsZWFwaXMuY29tlIwZX3VzZV9ub25fYmxvY2tpbmdfcmVmcmVzaJSJjAdfc2NvcGVzlF2UjCVodHRwczovL3d3dy5nb29nbGVhcGlzLmNvbS9hdXRoL2RyaXZllGGMD19kZWZhdWx0X3Njb3Blc5ROjA5fcmVmcmVzaF90b2tlbpSMZzEvLzBncEVOY2NoYmZCNXRDZ1lJQVJBQUdCQVNOd0YtTDlJcmtBYS1EajRBWm1pRVQwMGYyNVN3bE5VNU55MFo3X3ZLUEFXdi1oVnd0aXNNRXNBUDZDWGR0cWdLNnBseGNmenBKOHeUjAlfaWRfdG9rZW6UTowPX2dyYW50ZWRfc2NvcGVzlF2UjCVodHRwczovL3d3dy5nb29nbGVhcGlzLmNvbS9hdXRoL2RyaXZllGGMCl90b2tlbl91cmmUjCNodHRwczovL29hdXRoMi5nb29nbGVhcGlzLmNvbS90b2tlbpSMCl9jbGllbnRfaWSUjEg3Mzg1NzkzMzc0NTMtaTRiM2toYXA2ZjEwcmlybzhqOGM3ZmZqZnJoNGUzYzAuYXBwcy5nb29nbGV1c2VyY29udGVudC5jb22UjA5fY2xpZW50X3NlY3JldJSMI0dPQ1NQWC1jcnljX0JVWmM5VFlqMWNtRTNzajZVcmZHczZ6lIwLX3JhcHRfdG9rZW6UTowWX2VuYWJsZV9yZWF1dGhfcmVmcmVzaJSJjAhfYWNjb3VudJSMAJSMD19jcmVkX2ZpbGVfcGF0aJROdWIu"
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
    # 1. スプレッドシートをExcel形式でエクスポートして取得
    request_dl = service.files().export_media(
        fileId=EXCEL_FILE_ID,
        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request_dl)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)

    # 2. openpyxlで読み込み
    wb = load_workbook(fh)
    ws = wb.active

    # 3. 手動編集による「空行」対策：データが入っている本当の最終行を探す
    real_last_row = 0
    for row in range(ws.max_row, 0, -1):
        if any(cell.value is not None for cell in ws[row]):
            real_last_row = row
            break
    
    # 本当の最終行の次に追加
    new_row = real_last_row + 1
    ws.cell(row=new_row, column=1, value=filename)
    ws.cell(row=new_row, column=2, value=pay_date)
    ws.cell(row=new_row, column=3, value=payee)
    ws.cell(row=new_row, column=4, value=amount)

    # 4. メモリ上のバイナリに保存
    out_fh = io.BytesIO()
    wb.save(out_fh)
    out_fh.seek(0)

    # 5. 【重要】Googleドライブ上の既存ファイルを更新
    # スプレッドシート形式を維持したまま上書きする
    media = MediaIoBaseUpload(
        out_fh, 
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )
    # updateメソッドで、既存のファイルIDを指定して中身を入れ替える
    service.files().update(
        fileId=EXCEL_FILE_ID,
        media_body=media,
        supportsAllDrives=True
    ).execute()
    
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
