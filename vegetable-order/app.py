from flask import Flask, render_template, request, send_file
from docx import Document
from docx.table import Table
from datetime import datetime
import os

app = Flask(__name__)

TEMPLATE = os.path.join(os.path.dirname(__file__), "uploads", "form2026.docx")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    customer = request.form.get("customer", "")
    date = request.form.get("date", "")
    items = []

    # 网页粘贴菜品：每行一个菜品，只写入Word菜品列
    pasted = request.form.get("dishes", "")
    names = [x.strip() for x in pasted.splitlines() if x.strip()]

    for i in range(47):
        items.append({
            "dish": names[i] if i < len(names) else ""
        })

    doc = Document(TEMPLATE)

    # form2026实际由多个表格组成，python-docx默认只读取第一个表格
    # 获取文档中所有表格，只修改菜品编号对应的第一列
    all_tables = []
    for tbl in doc._body._body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"):
        all_tables.append(Table(tbl, doc._body))

    dish_cells = []
    # 第2个表格: 1-23，第3个表格: 24-47
    for t in all_tables[1:]:
        for r in t.rows:
            cells = r.cells
            if len(cells) >= 5:
                # 第一列是序号，第二列才是菜品
                dish_cells.append(cells[1])

    for idx, item in enumerate(items):
        if idx < len(dish_cells):
            dish_cells[idx].text = item["dish"]

    out = "output/蔬菜单_" + datetime.now().strftime("%Y%m%d%H%M%S") + ".docx"
    os.makedirs("output", exist_ok=True)
    doc.save(out)
    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(
        host="0.0.0.0",
        port=port
    )
