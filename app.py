import os
import io
import pickle
from flask import Flask, request, render_template
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from openpyxl import load_workbook

app = Flask(__name__)

# --- 環境変数 ---
CREDENTIALS_FILE = os.environ.get('GOOGLE_CREDENTIALS_JSON', 'credentials.json')
EXCEL_FILE_ID = os.environ['EXCEL_FILE_ID']  # Excel ファイルID
RECEIPTS_FOLDER_ID = os.environ['RECEIPTS_FOLDER_ID']  # 個人 Drive 内アップロード先フォルダID

SCOPES = ['https://www.googleapis.com/auth/drive.file']

# --- OAuth 認証 ---
def get_drive_service():
    creds = None
    # token.pickle があれば読み込む
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # なければ OAuth フロー
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    return service

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        drive_service = get_drive_service()

        file = request.files.get("file")
        date = request.form["date"]
        vendor = request.form["vendor"]
        amount = request.form["amount"]
        description = request.form["description"]

        # --- 画像ファイル名作成 ---
        filename = f"{vendor} {date} ¥{int(amount):,}.jpg"

        # --- Google Drive にアップロード ---
        media = MediaIoBaseUpload(file.stream, mimetype=file.mimetype)
        file_metadata = {
            "name": filename,
            "parents": [RECEIPTS_FOLDER_ID]
        }
        drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        # --- Excel ファイル更新 ---
        # Excel をダウンロード
        request_dl = drive_service.files().get_media(fileId=EXCEL_FILE_ID)
        fh = io.BytesIO(request_dl.execute())
        wb = load_workbook(fh)
        ws = wb.active
        ws.append([date, vendor, f"¥{int(amount):,}", description])
        out_fh = io.BytesIO()
        wb.save(out_fh)
        out_fh.seek(0)
        # 上書きアップロード
        drive_service.files().update(
            fileId=EXCEL_FILE_ID,
            media_body=MediaIoBaseUpload(out_fh, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        ).execute()

        return f"アップロード完了: {filename}"

    return render_template("index.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
