import os
import hashlib
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "regulation-lookup-secret"

BASE_DIR      = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
STATIC_IMG    = os.path.join(BASE_DIR, "static", "images")
ALLOWED_EXCEL = {"xlsx", "xls"}
ALLOWED_IMG   = {"png", "jpg", "jpeg", "gif", "webp"}
EXCEL_FILE    = os.path.join(UPLOAD_FOLDER, "data.xlsx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_IMG, exist_ok=True)

COL_DEFECT  = "缺失項目"
COL_REG     = "法源依據"
COL_CONTENT = "法規內容"

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_SS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# ─── helpers ─────────────────────────────────────────────────────────────────

def allowed_excel(f):
    return "." in f and f.rsplit(".", 1)[1].lower() in ALLOWED_EXCEL

def allowed_img(f):
    return "." in f and f.rsplit(".", 1)[1].lower() in ALLOWED_IMG

def get_excel_path():
    if os.path.exists(EXCEL_FILE):
        return EXCEL_FILE
    files = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith((".xlsx", ".xls"))]
    return os.path.join(UPLOAD_FOLDER, files[0]) if files else None

def find_col(columns, target):
    t = target.strip()
    for c in columns:
        if str(c).strip() == t:
            return c
    return None


# ─── xlsx image extraction ────────────────────────────────────────────────────

def _build_workbook_map(z):
    """
    回傳:
      name_to_sheetnum  : { sheet名稱: "N" }
      sheetnum_to_dnum  : { "N": "M" }  (有 drawing 的 sheet)
    """
    all_files = z.namelist()
    wb_xml  = ET.fromstring(z.read("xl/workbook.xml").decode())
    wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels").decode())

    rid_to_snum = {}
    for rel in wb_rels:
        t = rel.get("Target", ""); rid = rel.get("Id", "")
        if "sheet" in t.lower():
            rid_to_snum[rid] = t.replace("worksheets/sheet", "").replace(".xml", "")

    name_to_snum = {}
    for s in wb_xml.findall(f".//{{{NS_SS}}}sheet"):
        name = s.get("name", "")
        rid  = s.get(f"{{{NS_R}}}id", "")
        if rid in rid_to_snum:
            name_to_snum[name] = rid_to_snum[rid]

    snum_to_dnum = {}
    for snum in rid_to_snum.values():
        rp = f"xl/worksheets/_rels/sheet{snum}.xml.rels"
        if rp not in all_files:
            continue
        rr = ET.fromstring(z.read(rp).decode())
        for rel in rr:
            t = rel.get("Target", "")
            if "drawing" in t.lower():
                snum_to_dnum[snum] = t.split("drawing")[-1].replace(".xml", "")

    return name_to_snum, snum_to_dnum


def _build_richvalue_map(z):
    """
    解析 Excel IMAGE() 函數的 richData 格式。
    回傳 { vm_index: "xl/media/imageN.png" }
    vm_index 對應 cell 的 vm 屬性值。
    """
    all_files = z.namelist()
    rv_rel_path  = "xl/richData/richValueRel.xml"
    rv_rels_path = "xl/richData/_rels/richValueRel.xml.rels"
    rv_data_path = "xl/richData/rdrichvalue.xml"
    meta_path    = "xl/metadata.xml"

    if rv_rel_path not in all_files or rv_rels_path not in all_files:
        return {}

    # rId → media path
    rid_to_media = {}
    for rel in ET.fromstring(z.read(rv_rels_path).decode()):
        rid    = rel.get("Id", "")
        target = rel.get("Target", "")
        rid_to_media[rid] = "xl/" + target.replace("../", "")

    # richValueRel index (0-based) → rId
    NS_RVR = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
    rv_rel_root = ET.fromstring(z.read(rv_rel_path).decode())
    idx_to_rid = {}
    for i, rel in enumerate(rv_rel_root):
        rid = rel.get(f"{{{NS_R}}}id", "") or rel.get("r:id", "")
        idx_to_rid[i] = rid

    # rdrichvalue: rv entry N → LocalImageIdentifier (first v value)
    rv_local_ids = []
    if rv_data_path in all_files:
        NS_RD = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
        for rv in ET.fromstring(z.read(rv_data_path).decode()):
            vals = [v.text for v in rv.findall(f"{{{NS_RD}}}v")]
            rv_local_ids.append(int(vals[0]) if vals else 0)

    # metadata: futureMetadata index (= vm value) → rvb i → rv index → LocalImageIdentifier
    if meta_path not in all_files:
        return {}
    NS_META = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_XLRD = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
    meta_root = ET.fromstring(z.read(meta_path).decode())

    vm_to_media = {}
    fm = meta_root.find(f".//{{{NS_META}}}futureMetadata")
    if fm is None:
        return {}
    for vm_idx, bk in enumerate(fm.findall(f"{{{NS_META}}}bk")):
        rvb = bk.find(f".//{{{NS_XLRD}}}rvb")
        if rvb is None:
            continue
        rv_idx = int(rvb.get("i", 0))
        if rv_idx < len(rv_local_ids):
            local_id = rv_local_ids[rv_idx]
            rid = idx_to_rid.get(local_id, "")
            if rid in rid_to_media:
                vm_to_media[vm_idx] = rid_to_media[rid]
    return vm_to_media


def _scan_sheet_richvalue_cells(z, sheetnum):
    """
    掃描工作表 XML，找出所有含 vm 屬性的 cell（IMAGE() 函數）。
    回傳 [ { draw_row: int, vm: int } ]
    draw_row = Excel row - 1（draw_row 1 = 第一筆資料行）
    """
    sheet_path = f"xl/worksheets/sheet{sheetnum}.xml"
    if sheet_path not in z.namelist():
        return []
    NS_CELL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    root = ET.fromstring(z.read(sheet_path).decode())
    results = []
    for row_el in root.findall(f".//{{{NS_CELL}}}row"):
        for c in row_el.findall(f"{{{NS_CELL}}}c"):
            vm = c.get("vm")
            if vm is not None:
                ref = c.get("r", "")
                excel_row = int("".join(filter(str.isdigit, ref)))
                draw_row  = excel_row - 1   # draw_row 1 = Excel row 2 = first data row
                results.append({"draw_row": draw_row, "vm": int(vm)})
    return results


def _parse_drawing_images(z, d_num):
    """回傳 [ {draw_row: int, media: str} ]（drawing 格式，draw_row 已是 0-indexed）"""
    d_path = f"xl/drawings/drawing{d_num}.xml"
    d_rels = f"xl/drawings/_rels/drawing{d_num}.xml.rels"
    all_files = z.namelist()
    if d_path not in all_files:
        return []
    rid_to_media = {}
    if d_rels in all_files:
        for rel in ET.fromstring(z.read(d_rels).decode()):
            rid    = rel.get("Id", "")
            target = rel.get("Target", "")
            rid_to_media[rid] = "xl/" + target.replace("../", "")
    result = []
    root = ET.fromstring(z.read(d_path).decode())
    for anchor in root:
        tag = anchor.tag.split("}")[-1]
        if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue
        fe = anchor.find(f"{{{NS_XDR}}}from")
        row = 0
        if fe is not None:
            r_el = fe.find(f"{{{NS_XDR}}}row")
            row = int(r_el.text) if r_el is not None else 0
        blip = anchor.find(f".//{{{NS_A}}}blip")
        if blip is None:
            continue
        rid = blip.get(f"{{{NS_R}}}embed", "")
        if rid in rid_to_media:
            result.append({"draw_row": row, "media": rid_to_media[rid]})
    return result


def _save_image(img_bytes, ext, sheet_name, defect_idx):
    """存圖片到對應資料夾，回傳 filename 或 None"""
    if ext == "emf":
        return None   # EMF 格式瀏覽器無法顯示，跳過
    img_hash = hashlib.md5(img_bytes).hexdigest()[:10]
    filename = f"exc_{img_hash}.{ext}"
    folder   = _img_folder(sheet_name, defect_idx)
    filepath = os.path.join(folder, filename)
    if not os.path.exists(filepath):
        with open(filepath, "wb") as f:
            f.write(img_bytes)
    return filename


def _find_defect_idx(draw_row, row_starts):
    closest_idx = None
    min_dist = float("inf")

    for i, start in enumerate(row_starts):
        dist = abs(draw_row - start)

        if dist < min_dist:
            min_dist = dist
            closest_idx = i

    return closest_idx
    closest_idx = None
    min_dist = float("inf")

    for i, (s, e) in enumerate(defect_row_ranges):
    closest_idx = None
    min_dist = float("inf")

    for i, start in enumerate(row_starts):
        dist = abs(draw_row - start)

        if dist < min_dist:
            min_dist = dist
            closest_idx = i

    return closest_idx
    closest_idx = None
    min_dist = float("inf")

    for i, (s, e) in enumerate(defect_row_ranges):
        # 如果在範圍內 → 直接回傳
        if s <= draw_row <= e:
            return i

        # 計算距離（找最近的）
        dist = min(abs(draw_row - s), abs(draw_row - e))

        if dist < min_dist:
            min_dist = dist
            closest_idx = i

    return closest_idx
    for i, (s, e) in enumerate(defect_row_ranges):
        if s <= draw_row <= e:
            return i
    return None


def extract_and_cache_images(sheet_name, row_starts):
    path = get_excel_path()
    if not path:
        return {}

    result = {}

    try:
        with zipfile.ZipFile(path) as z:
            all_files = z.namelist()
            name_to_snum, snum_to_dnum = _build_workbook_map(z)
            sheetnum = name_to_snum.get(sheet_name)
            if sheetnum is None:
                return {}

            # ───── Drawing 圖片 ─────
            d_num = snum_to_dnum.get(sheetnum)
            if d_num:
                for img_info in _parse_drawing_images(z, d_num):
                    draw_row  = img_info["draw_row"]
                    media_key = img_info["media"]

                    if media_key not in all_files:
                        continue

                    defect_idx = _find_defect_idx(draw_row, row_starts)
                    if defect_idx is None:
                        continue

                    img_bytes = z.read(media_key)
                    ext = media_key.rsplit(".", 1)[-1].lower()
                    fname = _save_image(img_bytes, ext, sheet_name, defect_idx)

                    if fname:
                        result.setdefault(defect_idx, [])
                        result[defect_idx].append({
                            "row": draw_row,
                            "file": fname
                        })

            # ───── IMAGE() 圖片 ─────
            vm_to_media = _build_richvalue_map(z)
            if vm_to_media:
                for cell_info in _scan_sheet_richvalue_cells(z, sheetnum):
                    draw_row  = cell_info["draw_row"]
                    vm_idx    = cell_info["vm"]
                    media_key = vm_to_media.get(vm_idx)

                    if not media_key or media_key not in all_files:
                        continue

                    defect_idx = _find_defect_idx(draw_row, row_starts)
                    if defect_idx is None:
                        continue

                    img_bytes = z.read(media_key)
                    ext = media_key.rsplit(".", 1)[-1].lower()
                    fname = _save_image(img_bytes, ext, sheet_name, defect_idx)

                    if fname:
                        result.setdefault(defect_idx, [])
                        result[defect_idx].append({
                            "row": draw_row,
                            "file": fname
                        })

            # 🔥 排序（關鍵）
            for k in result:
                result[k] = [x["file"] for x in sorted(result[k], key=lambda x: x["row"])]

    except Exception as e:
        app.logger.error(f"extract_and_cache_images error: {e}")

    return result


def _sheet_dir(sheet_name):
    """用 hash 當資料夾名稱，避免中文被 secure_filename 全部清空而混用同一目錄"""
    return hashlib.md5(sheet_name.encode("utf-8")).hexdigest()[:12]


def _img_folder(sheet_name, item_index):
    folder = os.path.join(STATIC_IMG, _sheet_dir(sheet_name), str(item_index))
    os.makedirs(folder, exist_ok=True)
    return folder


def list_images(sheet_name, item_index):
    folder = _img_folder(sheet_name, item_index)
    return sorted([f for f in os.listdir(folder) if f.rsplit(".", 1)[-1].lower() in ALLOWED_IMG])


# ─── data loading ─────────────────────────────────────────────────────────────

def load_sheets():
    path = get_excel_path()
    if not path:
        return None, None
    try:
        return pd.ExcelFile(path).sheet_names, path
    except Exception as e:
        return None, str(e)


def load_sheet_data(sheet_name):
    path = get_excel_path()
    if not path:
        return None, [], []
    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=0)
        raw.columns = [str(c).strip() for c in raw.columns]
        actual_cols = list(raw.columns)

        col_defect = find_col(raw.columns, COL_DEFECT)
        col_reg    = find_col(raw.columns, COL_REG)
        if col_defect is None:
            return None, actual_cols, []

        records = []
        row_starts = []
        current = None

        for enum_idx, (_, row) in enumerate(raw.iterrows()):
            draw_row = enum_idx + 1

            val = str(row[col_defect]).strip() if pd.notna(row[col_defect]) else ""
            reg_val = str(row[col_reg]).strip() if col_reg and pd.notna(row[col_reg]) else ""

            if val and val not in ("nan", ""):
                if current is not None:
                    records.append(current)

                current = {
                    COL_DEFECT: val,
                    COL_REG: reg_val,
                    COL_CONTENT: ""
                }

                # 🔥 核心：只記「缺失開始的row」
                row_starts.append(draw_row)

            else:
                if current is not None and reg_val:
                    sep = "\n" if current[COL_CONTENT] else ""
                    current[COL_CONTENT] += sep + reg_val

        if current is not None:
            records.append(current)

        df = pd.DataFrame(records)

        # 🔥 注意：回傳 row_starts（不是 row_ranges）
        return df, actual_cols, row_starts

    except Exception as e:
        return None, [str(e)], []
    """
    Excel 結構「兩行一組」:
      行A: 項次, 缺失項目, 法源依據(短)
      行B: NaN,  NaN,      法規詳細內容
    同時回傳 defect_row_ranges: [(start_draw_row, end_draw_row), ...] (0-indexed)
    draw_row 0 = Excel header 行, draw_row 1 = 第一筆資料行
    """
    path = get_excel_path()
    if not path:
        return None, [], []
    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=0)
        raw.columns = [str(c).strip() for c in raw.columns]
        actual_cols = list(raw.columns)

        col_defect = find_col(raw.columns, COL_DEFECT)
        col_reg    = find_col(raw.columns, COL_REG)
        if col_defect is None:
            return None, actual_cols, []

        records     = []
        row_ranges  = []
        current     = None
        current_start = None

        for enum_idx, (_, row) in enumerate(raw.iterrows()):
            draw_row = enum_idx + 1   # +1 to skip header row (draw_row 0 = header)
            val     = str(row[col_defect]).strip() if pd.notna(row[col_defect]) else ""
            reg_val = str(row[col_reg]).strip()    if col_reg and pd.notna(row[col_reg]) else ""

            if val and val not in ("nan", ""):
                if current is not None:
                    row_ranges.append((current_start, draw_row - 1))
                    records.append(current)
                current       = {COL_DEFECT: val, COL_REG: reg_val, COL_CONTENT: ""}
                current_start = draw_row
            else:
                if current is not None and reg_val:
                    sep = "\n" if current[COL_CONTENT] else ""
                    current[COL_CONTENT] += sep + reg_val

        if current is not None:
            row_ranges.append((current_start, 99999))
            records.append(current)

        df = pd.DataFrame(records) if records else pd.DataFrame(columns=[COL_DEFECT, COL_REG, COL_CONTENT])
        return df, actual_cols, row_ranges

    except Exception as e:
        return None, [str(e)], []


# ─── routes ───────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or f.filename == "":
            flash("請選擇檔案")
        elif allowed_excel(f.filename):
            for old in os.listdir(UPLOAD_FOLDER):
                os.remove(os.path.join(UPLOAD_FOLDER, old))
            # 清除圖片快取
            import shutil
            if os.path.exists(STATIC_IMG):
                shutil.rmtree(STATIC_IMG)
            os.makedirs(STATIC_IMG, exist_ok=True)
            f.save(EXCEL_FILE)
            flash("檔案上傳成功！")
        else:
            flash("僅支援 .xlsx 或 .xls 格式")
        return redirect(url_for("index"))

    sheets, error = load_sheets()
    return render_template("index.html", sheets=sheets, error=error)


@app.route("/system/<path:sheet_name>")
def defects(sheet_name):
    df, _, _ = load_sheet_data(sheet_name)
    if df is None:
        return redirect(url_for("index"))
    items = df[COL_DEFECT].tolist()
    return render_template("defects.html", sheet_name=sheet_name, items=items)


@app.route("/system/<path:sheet_name>/defect/<int:item_index>", methods=["GET", "POST"])
def regulation(sheet_name, item_index):
    df, actual_cols, row_starts = load_sheet_data(sheet_name)
    if df is None or item_index >= len(df):
        return redirect(url_for("index"))

    # 手動上傳圖片
    if request.method == "POST":
        uploaded = request.files.getlist("images")
        count = 0
        for f in uploaded:
            if f and f.filename and allowed_img(f.filename):
                ext  = f.filename.rsplit(".", 1)[1].lower()
                name = hashlib.md5(f.read()).hexdigest()[:12] + "." + ext
                f.seek(0)
                f.save(os.path.join(_img_folder(sheet_name, item_index), name))
                count += 1
        if count:
            flash(f"上傳了 {count} 張圖片")
        return redirect(url_for("regulation", sheet_name=sheet_name, item_index=item_index))

    # 從 Excel 提取圖片（只提取這個工作表一次）
    extract_and_cache_images(sheet_name, row_starts)

    row          = df.iloc[item_index]
    defect       = row[COL_DEFECT]
    reg_text     = row[COL_REG]
    content_text = row[COL_CONTENT]
    images       = list_images(sheet_name, item_index)
    col_warning  = actual_cols if (not reg_text and not content_text and not images) else None

    return render_template(
        "regulation.html",
        sheet_name=sheet_name,
        defect=defect,
        reg_text=reg_text,
        content_text=content_text,
        images=images,
        item_index=item_index,
        col_warning=col_warning,
    )


@app.route("/system/<path:sheet_name>/defect/<int:item_index>/delete_image/<filename>")
def delete_image(sheet_name, item_index, filename):
    try:
        os.remove(os.path.join(_img_folder(sheet_name, item_index), secure_filename(filename)))
    except Exception:
        pass
    return redirect(url_for("regulation", sheet_name=sheet_name, item_index=item_index))


@app.route("/img/<path:sheet_name>/<int:item_index>/<filename>")
def serve_image(sheet_name, item_index, filename):
    return send_from_directory(_img_folder(sheet_name, item_index), filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
