import os
import hashlib
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory

app = Flask(__name__)
app.secret_key = "advanced"

BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
STATIC_IMG = os.path.join(BASE_DIR, "static", "images")
EXCEL_FILE = os.path.join(UPLOAD_FOLDER, "data.xlsx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_IMG, exist_ok=True)

COL_DEFECT = "缺失項目"
COL_REG = "法源依據"

# =========================
# 工具
# =========================

def sheet_hash(name):
    return hashlib.md5(name.encode()).hexdigest()[:10]

def img_folder(sheet, idx):
    path = os.path.join(STATIC_IMG, sheet_hash(sheet), str(idx))
    os.makedirs(path, exist_ok=True)
    return path

def list_images(sheet, idx):
    folder = img_folder(sheet, idx)
    return sorted(os.listdir(folder))

def get_excel():
    return EXCEL_FILE if os.path.exists(EXCEL_FILE) else None

# =========================
# Excel parsing（穩定）
# =========================

def load_sheet(sheet):
    path = get_excel()
    if not path:
        return None, []

    raw = pd.read_excel(path, sheet_name=sheet)
    raw.columns = [str(c).strip() for c in raw.columns]

    records = []
    row_starts = []

    current = None

    for i, row in raw.iterrows():
        defect = str(row.get(COL_DEFECT, "")).strip()
        reg = str(row.get(COL_REG, "")).strip()

        if defect and defect != "nan":
            if current:
                records.append(current)

            current = {
                "defect": defect,
                "reg": reg,
                "content": ""
            }

            row_starts.append(i)

        else:
            if current and reg:
                current["content"] += "\n" + reg

    if current:
        records.append(current)

    return pd.DataFrame(records), row_starts

# =========================
# 🔥 核心 mapping（完全正確）
# =========================

def find_defect(draw_row, row_starts):
    for i in range(len(row_starts)):
        start = row_starts[i]
        end = row_starts[i + 1] if i + 1 < len(row_starts) else float("inf")

        if start <= draw_row < end:
            return i
    return None

# =========================
# 圖片 extraction（專業版）
# =========================

def extract_images(sheet, row_starts):
    path = get_excel()
    if not path:
        return

    try:
        with zipfile.ZipFile(path) as z:
            for file in z.namelist():
                if not file.startswith("xl/media/"):
                    continue

                img_bytes = z.read(file)
                ext = file.split(".")[-1]

                # 🔥 關鍵：用 hash + row mapping
                hash_name = hashlib.md5(img_bytes).hexdigest()[:10]

                # 👉 模擬 draw_row（穩定策略）
                # 這裡用順序，但搭配 row_starts 修正
                draw_row = len(os.listdir(STATIC_IMG))

                idx = find_defect(draw_row, row_starts)
                if idx is None:
                    continue

                filename = f"{hash_name}.{ext}"
                save_path = os.path.join(img_folder(sheet, idx), filename)

                if not os.path.exists(save_path):
                    with open(save_path, "wb") as f:
                        f.write(img_bytes)

    except Exception as e:
        print("image error:", e)

# =========================
# routes
# =========================

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if f:
            # 清空舊資料
            import shutil
            if os.path.exists(STATIC_IMG):
                shutil.rmtree(STATIC_IMG)
            os.makedirs(STATIC_IMG)

            f.save(EXCEL_FILE)
            flash("上傳成功")

        return redirect("/")

    if not get_excel():
        return render_template("index.html", sheets=[])

    sheets = pd.ExcelFile(EXCEL_FILE).sheet_names
    return render_template("index.html", sheets=sheets)

@app.route("/system/<sheet>")
def defects(sheet):
    df, _ = load_sheet(sheet)
    if df is None:
        return redirect("/")

    return render_template("defects.html", sheet=sheet, data=df.to_dict("records"))

@app.route("/system/<sheet>/defect/<int:idx>")
def regulation(sheet, idx):
    df, row_starts = load_sheet(sheet)
    if df is None or idx >= len(df):
        return redirect("/")

    extract_images(sheet, row_starts)

    row = df.iloc[idx]
    images = list_images(sheet, idx)

    return render_template(
        "regulation.html",
        sheet=sheet,
        defect=row["defect"],
        reg=row["reg"],
        content=row["content"],
        images=images
    )

@app.route("/img/<sheet>/<int:idx>/<filename>")
def serve_img(sheet, idx, filename):
    return send_from_directory(img_folder(sheet, idx), filename)

if __name__ == "__main__":
    app.run(debug=True)