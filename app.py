import os
import io
import time
from typing import List, Dict, Any
from flask import Flask, request, render_template_string, redirect, url_for, flash
from werkzeug.utils import secure_filename
import pandas as pd

# -------------------- Config --------------------
UPLOAD_DIR = "uploads"
ALLOWED_EXT = {".xlsx", ".xls", ".csv"}
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = "change-this-in-production"  # เปลี่ยนซะถ้าจะใช้นานๆ

# เก็บไฟล์ที่อัปโหลดไว้ในหน่วยความจำ (ง่ายและเร็ว)
# โครงสร้าง: {"filename": str, "sheet": str, "df": pd.DataFrame, "path": str}
DATASTORE: List[Dict[str, Any]] = []

# -------------------- Utils --------------------
def ext_of(filename: str) -> str:
    return os.path.splitext(filename)[1].lower()

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
            mask = df_str[c].str.contains(kw, case=False, regex=False, na=False)
            idxs = df.index[mask].tolist()
            for idx in idxs:
                val = df_str.at[idx, c]
                # คำนวณหมายเลขแถวแบบ Excel: header = 1, data row แรก = 2
                # ถ้า index เป็น 0..N-1 ก็ใช้ idx + 2 ได้เลย
                try:
                    excel_row = int(idx) + 2
                    data_row = int(idx) + 1
                except Exception:
                    # ถ้าเป็น index แบบ label แปลกๆ รีเซ็ตนับเอง
                    # หาอันดับที่เท่าไหร่ในตำแหน่งที่แมตช์
                    pos = mask[mask].index.get_loc(idx) if hasattr(mask.index, "get_loc") else 0
                    excel_row = pos + 2
                    data_row = pos + 1

                results.append({
                    "filename": item["filename"],
                    "sheet": item["sheet"],
                    "excel_row": excel_row,
                    "data_row": data_row,
                    "column": str(c),
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
    <h1>🔎 ค้นหาชื่อใน Excel/CSV</h1>
    <form action="{{ url_for('clear') }}" method="post">
      <button type="submit" title="ล้างรายการไฟล์ทั้งหมด">ล้างไฟล์ทั้งหมด</button>
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
    <h3>1) อัปโหลดไฟล์ (.xlsx / .csv) จะอัปหลายไฟล์ก็ได้</h3>
    <form action="{{ url_for('upload') }}" method="post" enctype="multipart/form-data">
      <input type="file" name="files" multiple accept=".xlsx,.xls,.csv" required>
      <button type="submit">อัปโหลด</button>
    </form>
    <p class="muted">ไฟล์ที่รองรับหลัก: .xlsx, .csv  ถ้า .xls ต้องติดตั้ง xlrd เพิ่ม</p>
  </div>

  <div class="card">
    <h3>2) ค้นหา</h3>
    <form action="{{ url_for('search') }}" method="post" class="row">
      <div class="col">
        <label>คำค้นหา</label>
        <input type="text" name="q" placeholder="พิมพ์ชื่อหรือบางส่วนของชื่อ" required>
      </div>
      <div class="col" style="align-self:flex-end">
        <button type="submit">ค้นหา</button>
      </div>
    </form>
  </div>

  <div class="card">
    <h3>ไฟล์ที่อัปโหลดแล้ว ({{ files|length }} ไฟล์)</h3>
    {% if files %}
      <ul class="files">
        {% for f in files %}
          <li>• {{ f.filename }}{% if f.sheet and f.sheet != '-' %} (ชีต: {{ f.sheet }}){% endif %}</li>
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

@app.route("/", methods=["GET"])
def index():
    return render_template_string(INDEX_HTML, files=DATASTORE, results=None, query=None)

@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("files")
    if not files:
        flash("กรุณาเลือกไฟล์ก่อน")
        return redirect(url_for("index"))

    added = 0
    for f in files:
        if not f or f.filename == "":
            continue
        filename = secure_filename(f.filename)
        ext = ext_of(filename)
        if ext not in ALLOWED_EXT:
            flash(f"ข้ามไฟล์ {filename}: นามสกุลไม่รองรับ")
            continue

        # ตั้งชื่อกันชนกัน
        unique = f"{int(time.time()*1000)}_{filename}"
        path = os.path.join(UPLOAD_DIR, unique)
        f.save(path)

        try:
            item = load_dataframe_from_file(path, filename)
            DATASTORE.append(item)
            added += 1
        except Exception as e:
            flash(f"อ่านไฟล์ {filename} ไม่สำเร็จ: {e}")
            try:
                os.remove(path)
            except Exception:
                pass

    flash(f"อัปโหลดสำเร็จ {added} ไฟล์")
    return redirect(url_for("index"))

@app.route("/search", methods=["POST"])
def search():
    q = request.form.get("q", "").strip()
    if not q:
      flash("ใส่คำค้นหาก่อนสิ")
      return redirect(url_for("index"))

    if not DATASTORE:
      flash("ยังไม่มีไฟล์ให้อวดผล ควรอัปโหลดก่อน")
      return redirect(url_for("index"))

    results = search_in_datastore(q)
    return render_template_string(INDEX_HTML, files=DATASTORE, results=results, query=q)

@app.route("/clear", methods=["POST"])
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
