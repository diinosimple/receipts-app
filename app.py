from flask import Flask, request
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
EXCEL_FILE = "receipts.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Excelがなければ作成
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append(["支払日", "支払先", "内容", "画像"])
    wb.save(EXCEL_FILE)

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
    image.save(os.path.join(UPLOAD_FOLDER, filename))

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([pay_date, pay_to, description, filename])
    wb.save(EXCEL_FILE)

    return "保存しました"
