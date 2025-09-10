import os
import io
import time
from typing import List, Dict, Any
from flask import Flask, request, render_template_string, redirect, url_for, flash
from werkzeug.utils import secure_filename
import pandas as pd

# -------------------- Config --------------------
UPLOAD_DIR = "uploads"
ALLOWED_EXT = {".xlsx", ".xls", ".csv", ".txt"}  # เพิ่ม .txt
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = "change-this-in-production"  # เปลี่ยนซะถ้าจะใช้นานๆ

# เก็บไฟล์ที่อัปโหลดไว้ในหน่วยความจำ (ง่ายและเร็ว)
# โครงสร้าง: {"filename": str, "sheet": str, "df": pd.DataFrame, "path": str}
DATASTORE: List[Dict[str, Any]] = []

# -------------------- Utils --------------------
def ext_of(filename: str) -> str:
    return os.path.splitext(filename)[1].lower()


def _load_txt_as_dataframe(path: str, filename: str) -> Dict[str, Any]:
    """
    อ่านไฟล์ .txt แล้วแปลงเป็น DataFrame หนึ่งคอลัมน์ชื่อ 'text'
    บรรทัดละ 1 แถว เพื่อให้ search pipeline เดิมทำงานได้ทันที
    """
    encodings = ("utf-8", "utf-8-sig", "cp874", "iso-8859-11", "cp1252")
    with open(path, "rb") as f:
        raw = f.read()
    text = None
    for enc in encodings:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")
    # normalize newlines
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    # ถ้าไม่อยากเก็บบรรทัดว่าง คอมเมนต์บรรทัดด้านล่างออก
    # lines = [ln for ln in lines if ln.strip()]
    df = pd.DataFrame({"text": lines})
    return {"filename": filename, "sheet": "-", "df": df, "path": path}


def load_dataframe_from_file(path: str, filename: str) -> Dict[str, Any]:
    ext = ext_of(filename)
    if ext in {".xlsx", ".xls"}:
        # อ่านชีตแรกพอ (ส่วนใหญ่พอแล้ว)
        try:
            df = pd.read_excel(path)
            sheet = "Sheet1"
        except Exception:
            # บางไฟล์ระบุชีตแปลกๆ ลองอ่านทั้งหมดแล้วเอาอันแรก
            xls = pd.ExcelFile(path)
            first_sheet = xls.sheet_names[0]
            df = pd.read_excel(path, sheet_name=first_sheet)
            sheet = first_sheet
        return {"filename": filename, "sheet": sheet, "df": df, "path": path}
    elif ext == ".csv":
        # เดา encoding แบบบ้านๆ เผื่อมี BOM/ไทย
        with open(path, "rb") as f:
            raw = f.read()
        try:
            df = pd.read_csv(io.BytesIO(raw))
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig")
        return {"filename": filename, "sheet": "-", "df": df, "path": path}
    elif ext == ".txt":
        # โหลด .txt เป็น DataFrame หนึ่งคอลัมน์ชื่อ 'text'
        return _load_txt_as_dataframe(path, filename)
    else:
        raise ValueError("นามสกุลไฟล์ไม่รองรับ")


def search_in_datastore(keyword: str) -> List[Dict[str, Any]]:
    """
    ค้นหาแบบไม่แยกเล็กใหญ่ ในทุกคอลัมน์ที่เป็นข้อความ
    คืนรายการผลลัพธ์ราย cell: filename, sheet, excel_row, data_row, column, value
    - excel_row นับแบบ Excel โดยคิดว่าบรรทัดที่ 1 คือ header ดังนั้นแถวข้อมูลบรรทัดแรกคือ 2
    - data_row เป็นลำดับข้อมูล 1-based (ไม่นับ header)
    """
    kw = str(keyword)
    results = []
    for item in DATASTORE:
        df = item["df"]
        # ระบุคอลัมน์ข้อความ
        text_cols = [
            c for c in df.columns
            if pd.api.types.is_string_dtype(df[c]) or pd.api.types.is_object_dtype(df[c])
        ]
        if not text_cols:
            continue

        # แปลงเป็นสตริงกัน NaN
        df_str = df.copy()
        for c in text_cols:
            df_str[c] = df_str[c].astype(str)

        # ค้นหาทีละคอลัมน์
        for c in text_cols:
            mask = df_str[c].str.contains(kw, case=False, na=False)
            idxs = df_str.index[mask].tolist()
            for idx in idxs:
                # excel_row: header อยู่แถว 1, ข้อมูลเริ่มแถว 2
                excel_row = idx + 2
                data_row = idx + 1
                val = df_str.at[idx, c]
                results.append({
                    "filename": item["filename"],
                    "sheet": item["sheet"],
                    "excel_row": excel_row,
                    "data_row": data_row,
                    "column": c,
                    "value": val
                })
    return results

# -------------------- Routes --------------------
INDEX_HTML = """
<!doctype html>
<html lang="th">
<head>
  <meta charset="utf-8">
  <title>ค้นหาชื่อใน Excel/CSV</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif; padding:24px; max-width: 1100px; margin: auto;}
    header{display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap;}
    .card{border:1px solid #ddd; border-radius:12px; padding:16px; margin:16px 0;}
    .muted{color:#666; font-size:14px;}
    .files li{margin:4px 0;}
    input[type="text"]{padding:10px; font-size:16px; width:100%; max-width:400px;}
    input[type="file"]{padding:8px;}
    button{padding:10px 14px; font-size:15px; cursor:pointer; border-radius:10px; border:1px solid #ccc; background:#f7f7f7;}
    table{border-collapse:collapse; width:100%;}
    th, td{border:1px solid #e5e5e5; padding:8px; text-align:left; font-size:14px;}
    th{background:#fafafa;}
    .ok{color: #0b7; font-weight: 600;}
    .bad{color: #b00; font-weight: 600;}
    .row{display:flex; gap:16px; flex-wrap:wrap;}
    .col{flex:1 1 300px;}
    .flash{padding:10px 14px; border:1px solid #ffd18a; background:#fff5e6; border-radius:10px; margin:8px 0;}
  </style>
</head>
<body>
  <header>
    <h1>ค้นหาชื่อในไฟล์ Excel/CSV/TXT</h1>
    <form action="{{ url_for('clear') }}" method="post">
      <button type="submit">ล้างไฟล์ทั้งหมด</button>
    </form>
  </header>

  {% with messages = get_flashed_messages() %}
    {% if messages %}
      {% for m in messages %}
        <div class="flash">{{ m }}</div>
      {% endfor %}
    {% endif %}
  {% endwith %}

  <div class="card">
    <h3>อัปโหลดไฟล์ (.xlsx, .xls, .csv, .txt)</h3>
    <form action="{{ url_for('upload') }}" method="post" enctype="multipart/form-data">
      <input type="file" name="file" required>
      <button type="submit">อัปโหลด</button>
      <p class="muted">ไฟล์จะถูกเก็บในเครื่องเพื่อค้นหา และคงอยู่จนกว่าจะกดล้าง</p>
    </form>
  </div>

  <div class="card">
    <h3>ค้นหา</h3>
    <form action="{{ url_for('index') }}" method="get" class="row">
      <div class="col">
        <input type="text" name="q" placeholder="พิมพ์คำที่ต้องการค้นหา" value="{{ query or '' }}">
      </div>
      <div class="col">
        <button type="submit">ค้นหา</button>
      </div>
    </form>
  </div>

  <div class="card">
    <h3>ไฟล์ที่มีในระบบ</h3>
    {% if files %}
      <ul class="files">
        {% for f in files %}
          <li>{{ f }}</li>
        {% endfor %}
      </ul>
    {% else %}
      <p class="muted">ยังไม่มีไฟล์ ใส่มาก่อนค่อยอวด</p>
    {% endif %}
  </div>

  {% if query is not none %}
    <div class="card">
      <h3>ผลการค้นหา: "{{ query }}"</h3>
      {% if results %}
        <p class="ok">พบทั้งหมด {{ results|length }} ตำแหน่ง</p>
        <table>
          <thead>
            <tr>
              <th>ไฟล์</th>
              <th>ค่าในเซลล์</th>
              <th>แถว (ข้อมูล)</th>
            </tr>
          </thead>
          <tbody>
            {% for r in results %}
              <tr>
                <td>{{ r.filename }}</td>
                <td>{{ r.value }}</td>
                <td>{{ r.data_row }}</td>
              </tr>
            {% endfor %}
          </tbody>
        </table>
      {% else %}
        <p class="bad">ไม่พบผลลัพธ์ในไฟล์ที่อัปโหลด</p>
      {% endif %}
    </div>
  {% endif %}
</body>
</html>
"""

@app.get("/")
def index():
    query = request.args.get("q")
    results = None
    if query is not None:
        results = search_in_datastore(query)

    files = [item["filename"] for item in DATASTORE]
    return render_template_string(INDEX_HTML, query=query, results=results, files=files)


@app.post("/upload")
def upload():
    f = request.files.get("file")
    if not f or f.filename == "":
        flash("ไม่พบไฟล์ที่จะอัปโหลด")
        return redirect(url_for("index"))

    filename = secure_filename(f.filename)
    ext = ext_of(filename)
    if ext not in ALLOWED_EXT:
        flash(f"นามสกุลไม่รองรับ: {ext}")
        return redirect(url_for("index"))

    # บันทึกไฟล์ลงดิสก์ (Render ให้ mount disk ได้ ถ้าอยากเก็บรอดดีพลอย)
    save_path = os.path.join(UPLOAD_DIR, filename)
    f.save(save_path)

    try:
        item = load_dataframe_from_file(save_path, filename)
    except Exception as e:
        flash(f"โหลดไฟล์ไม่สำเร็จ: {e}")
        try:
            os.remove(save_path)
        except Exception:
            pass
        return redirect(url_for("index"))

    DATASTORE.append(item)
    flash(f"อัปโหลดแล้ว: {filename} (แถวข้อมูล {len(item['df'])})")
    return redirect(url_for("index"))


@app.post("/clear")
def clear():
    # ล้างรายการในหน่วยความจำ และลบไฟล์ที่อัปโหลด
    for item in DATASTORE:
        try:
            if os.path.exists(item["path"]):
                os.remove(item["path"])
        except Exception:
            pass
    DATASTORE.clear()
    flash("ล้างไฟล์ทั้งหมดแล้ว")
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)
