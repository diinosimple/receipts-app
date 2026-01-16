from flask import Flask, request
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os, io, pickle

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

app = Flask(__name__)
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

BASE_FOLDER = "ReceiptsApp"
IMAGES_FOLDER = "images"
EXCEL_NAME = "receipts.xlsx"

# -----------------------
# Google認証
# -----------------------
def get_drive_service():
    creds = None
    if os.path.exists("token.pickle"):
        creds = pickle.load(open("token.pickle", "rb"))

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        pickle.dump(creds, open("token.pickle", "wb"))

    return build("drive", "v3", credentials=creds)

service = get_drive_service()

# -----------------------
# Drive操作
# -----------------------
def get_or_create_folder(name, parent=None):
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder'"
    if parent:
        q += f" and '{parent}' in parents"
    res = service.files().list(q=q, fields="files(id)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]

    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    if parent:
        meta["parents"] = [parent]
    folder = service.files().create(body=meta, fields="id").execute()
    return folder["id"]

BASE_ID = get_or_create_folder(BASE_FOLDER)
IMAGES_ID = get_or_create_folder(IMAGES_FOLDER, BASE_ID)

def get_or_create_excel():
    q = f"name='{EXCEL_NAME}' and '{BASE_ID}' in parents"
    res = service.files().list(q=q, fields="files(id)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]

    wb = Workbook()
    ws = wb.active
    ws.append(["支払日", "支払先", "内容", "画像URL"])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    media = MediaIoBaseUpload(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    meta = {"name": EXCEL_NAME, "parents": [BASE_ID]}
    file = service.files().create(body=meta, media_body=media, fields="id").execute()
    return file["id"]

EXCEL_ID = get_or_create_excel()

# Excelダウンロード
def download_excel():
    buf = io.BytesIO()
    service.files().get_media(fileId=EXCEL_ID).execute()
    request = service.files().get_media(fileId=EXCEL_ID)
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    buf.seek(0)
    return buf

# Excelアップロード
def upload_excel(buf):
    media = MediaIoBaseUpload(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    service.files().update(fileId=EXCEL_ID, media_body=media).execute()

# -----------------------
# Web
# -----------------------
@app.route("/")
def index():
    return '''
    <form method="POST" action="/upload" enctype="multipart/form-data">
      <input type="date" name="pay_date" required><br>
      <input type="text" name="pay_to" placeholder="支払先" required><br>
      <input type="text" name="description" placeholder="内容" required><br>
      <input type="file" name="image" accept="image/*" capture="camera" required><br>
      <button type="submit">送信</button>
    </form>
    '''

@app.route("/upload", methods=["POST"])
def upload():
    image = request.files["image"]
    pay_date = request.form["pay_date"]
    pay_to = request.form["pay_to"]
    description = request.form["description"]

    filename = datetime.now().strftime("%Y%m%d_%H%M%S_") + image.filename

    # 画像をDriveへ
    media = MediaIoBaseUpload(image.stream, mimetype=image.mimetype)
    meta = {"name": filename, "parents": [IMAGES_ID]}
    file = service.files().create(body=meta, media_body=media, fields="id").execute()

    file_id = file["id"]

    service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()

    url = f"https://drive.google.com/uc?id={file_id}"

    # Excel更新
    buf = download_excel()
    wb = load_workbook(buf)
    ws = wb.active
    ws.append([pay_date, pay_to, description, url])

    new_buf = io.BytesIO()
    wb.save(new_buf)
    new_buf.seek(0)
    upload_excel(new_buf)

    return "保存しました"
