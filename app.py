import os
import io
import datetime
import pickle

from flask import Flask, request, redirect, url_for, render_template_string
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from openpyxl import load_workbook

# ========================
# Flask
# ========================
app = Flask(__name__)

# ========================
# Google Drive 認証
# ========================
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

def get_drive_service():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("token.pickle not found or invalid")

    return build("drive", "v3", credentials=creds)

# ========================
# Drive 操作
# ========================
def find_folder(service, name, parent_id=None):
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder'"
    if parent_id:
        q += f" and '{parent_id}' in parents"

    res = service.files().list(q=q, spaces="drive").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def find_file(service, name, parent_id):
    q = f"name='{name}' and '{parent_id}' in parents"
    res = service.files().list(q=q, spaces="drive").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


# ========================
# 画面
# ========================
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
    <input
      type="file"
      name="image"
      accept="image/*"
      capture="environment"
      required
    >
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

  <p>
    <button type="submit">アップロード</button>
  </p>
</form>

</body>
</html>
"""



# ========================
# ルーティング
# ========================
@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        service = get_drive_service()

        # フォルダ取得
        root_id = find_folder(service, "ReceiptsApp")
        images_id = find_folder(service, "images", root_id)

        if not root_id or not images_id:
            return "Drive フォルダが見つかりません"

        # ===== 画像保存 =====
        image = request.files["image"]
        filename = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_") + image.filename

        media = MediaIoBaseUpload(
            io.BytesIO(image.read()),
            mimetype=image.content_type,
            resumable=False,
        )

        file_metadata = {
            "name": filename,
            "parents": [images_id],
        }

        service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id",
        ).execute()

        # ===== Excel 更新 =====
        excel_id = find_file(service, "receipts.xlsx", root_id)
        if not excel_id:
            return "receipts.xlsx が見つかりません"

        # ダウンロード
        fh = io.BytesIO()
        request_dl = service.files().get_media(fileId=excel_id)
        downloader = MediaIoBaseDownload(fh, request_dl)
        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)
        wb = load_workbook(fh)
        ws = wb.active

        ws.append([
            request.form["date"],
            request.form["payee"],
            request.form["description"],
            filename,
        ])

        # アップロード（上書き）
        fh_out = io.BytesIO()
        wb.save(fh_out)
        fh_out.seek(0)

        media_excel = MediaIoBaseUpload(
            fh_out,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False,
        )

        service.files().update(
            fileId=excel_id,
            media_body=media_excel,
        ).execute()

        return redirect(url_for("upload"))

    return render_template_string(HTML)

# ========================
# 起動
# ========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=True)
