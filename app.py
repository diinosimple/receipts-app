import os
import io
import base64
import pickle

from flask import Flask, request
from werkzeug.utils import secure_filename

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.credentials import Credentials

from openpyxl import load_workbook
from openpyxl import Workbook

app = Flask(__name__)

# ===== 設定 =====
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
RECEIPTS_FOLDER_ID = "1YzEouvialQDhMEkWPpjG-elvD7Gb3Csy"
EXCEL_FILE_NAME = "receipts.xlsx"


# ===== HTML =====
HTML = """
<!doctype html>
<html>
<head>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Receipt Upload</title>
</head>
<body>
<h2>領収書アップロード</h2>

<form method="post" enctype="multipart/form-data">
  <p>
    <input type="file" name="image" accept="image/*" capture="environment" required>
  </p>

  <p>支払日:<br>
     <input type="date" name="date" required>
  </p>

  <p>支払い先:<br>
     <input type="text" name="payee" required>
  </p>

  <p>内容:<br>
     <input type="text" name="description" required>
  </p>

  <p>金額:<br>
     <input type="number" name="amount" required>
  </p>

  <p>
    <button type="submit">アップロード</button>
  </p>
</form>

</body>
</html>
"""


# ===== token.pickle 復元（Render用）=====
def restore_token_from_env():
    if os.path.exists("token.pickle"):
        return

    token_b64 = os.environ.get("TOKEN_PICKLE_BASE64")
    if not token_b64:
        return

    with open("token.pickle", "wb") as f:
        f.write(base64.b64decode(token_b64))


# ===== Google Drive Service =====
def get_drive_service():
    restore_token_from_env()

    with open("token.pickle", "rb") as f:
        creds = pickle.load(f)

    return build("drive", "v3", credentials=creds)


# ===== Excel ファイル取得 or 作成 =====

def get_or_create_excel(service):
    query = f"name='{EXCEL_FILE_NAME}' and '{RECEIPTS_FOLDER_ID}' in parents"
    res = service.files().list(q=query, fields="files(id,name)").execute()
    files = res.get("files", [])

    if files:
        return files[0]["id"]

    # 新規作成
    wb = Workbook()
    ws = wb.active
    ws.append(["支払日", "支払い先", "内容", "金額"])

    buf = io.BytesIO()
    wb.save(buf)

    # ★これが超重要
    buf.seek(0)

    media = MediaIoBaseUpload(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False
    )

    file_metadata = {
        "name": EXCEL_FILE_NAME,
        "parents": [RECEIPTS_FOLDER_ID]
    }

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    return file["id"]


# ===== Excel に追記 =====
def append_excel(service, excel_id, row):
    request = service.files().get_media(fileId=excel_id)
    fh = io.BytesIO(request.execute())

    wb = load_workbook(fh)
    ws = wb.active
    ws.append(row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    media = MediaIoBaseUpload(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False
    )

    service.files().update(
        fileId=excel_id,
        media_body=media
    ).execute()


# ===== メイン =====
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return HTML

    service = get_drive_service()

    # フォーム値取得
    date = request.form["date"]
    payee = request.form["payee"]
    description = request.form["description"]

    amount_int = int(request.form["amount"])
    amount_display = f"¥{amount_int:,}"
    amount_for_filename = f"Y{amount_int}"

    # ファイル名生成
    base_name = f"{payee} {date} {amount_for_filename}"
    safe_name = secure_filename(base_name)
    filename = f"{safe_name}.jpg"

    # 画像アップロード
    image = request.files["image"]
    media = MediaIoBaseUpload(image.stream, mimetype="image/jpeg")

    service.files().create(
        body={
            "name": filename,
            "parents": [RECEIPTS_FOLDER_ID]
        },
        media_body=media
    ).execute()

    # Excel 追記
    excel_id = get_or_create_excel(service)
    append_excel(service, excel_id, [date, payee, description, amount_display])

    return "アップロード完了しました 👍"


# ===== 起動 =====
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=True)