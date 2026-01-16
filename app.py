import os
import io
import json
import pickle
from flask import Flask, request, render_template
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from openpyxl import load_workbook

app = Flask(__name__)

# --- 環境変数 ---
SCOPES = ['https://www.googleapis.com/auth/drive.file']
EXCEL_FILE_ID = os.environ['EXCEL_FILE_ID']         # Excel ファイルID
RECEIPTS_FOLDER_ID = os.environ['RECEIPTS_FOLDER_ID']  # 個人 Drive フォルダID
GOOGLE_CREDENTIALS_JSON = os.environ['GOOGLE_CREDENTIALS_JSON']  # OAuth クライアント情報

TOKEN_FILE = 'token.pickle'


# --- Drive サービス取得 ---
def get_drive_service():
    creds = None
    # 既存 token があれば読み込む
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)

    # 新規認証が必要な場合
    if not creds or not creds.valid:
        client_config = json.loads(GOOGLE_CREDENTIALS_JSON)
        flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service


# --- Flask ルート ---
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

        # --- Drive にアップロード ---
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

        # --- Excel 更新 ---
        request_dl = drive_service.files().get_media(fileId=EXCEL_FILE_ID)
        fh = io.BytesIO(request_dl.execute())
        wb = load_workbook(fh)
        ws = wb.active
        ws.append([date, vendor, f"¥{int(amount):,}", description])
        out_fh = io.BytesIO()
        wb.save(out_fh)
        out_fh.seek(0)
        drive_service.files().update(
            fileId=EXCEL_FILE_ID,
            media_body=MediaIoBaseUpload(out_fh, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        ).execute()

        return f"アップロード完了: {filename}"

    # GET は templates/index.html をレンダリング
    return render_template("index.html")


# --- メイン ---
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)