from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

import os
import datetime
import tempfile


app = Flask(__name__)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PATH = os.path.join(
    BASE_DIR,
    "uploads",
    "form2026.docx"
)



def set_font(run, size, bold=True):
    """
    设置字体
    """
    run.font.name = "宋体"
    run.font.size = Pt(size)
    run.bold = bold



def write_customer_and_date(doc, customer):

    today = datetime.datetime.now().strftime(
        "%Y年%m月%d日"
    )


    # 遍历所有段落寻找客户和日期位置

    for p in doc.paragraphs:

        text = p.text


        if "客户：" in text:

            p.clear()

            run = p.add_run(
                "客户：" + customer
            )

            set_font(
                run,
                14,
                True
            )


        elif "2026年" in text:

            p.clear()

            run = p.add_run(
                today
            )

            set_font(
                run,
                14,
                True
            )



def write_dishes(doc, dishes):

    count = 0


    # form2026 有多个表格
    # 遍历所有表格寻找数字序号行

    for table in doc.tables:

        for row in table.rows:

            cells = row.cells


            if len(cells) < 4:
                continue


            first = cells[0].text.strip()


            # 找 1-47 行
            if first.isdigit():

                num = int(first)


                if 1 <= num <= 47:

                    if count < len(dishes):

                        # 只修改第一列

                        cell = cells[0]


                        cell.text = ""


                        p = cell.paragraphs[0]


                        run = p.add_run(
                            str(num)
                            + "  "
                            + dishes[count]
                        )


                        set_font(
                            run,
                            18,
                            True
                        )


                        count += 1



def create_word(customer, dishes):


    doc = Document(
        TEMPLATE_PATH
    )


    # 写客户日期

    write_customer_and_date(
        doc,
        customer
    )


    # 写菜品

    write_dishes(
        doc,
        dishes
    )



    filename = (
        "蔬菜订单_"
        +
        datetime.datetime.now()
        .strftime("%Y%m%d%H%M%S")
        +
        ".docx"
    )


    output = os.path.join(
        tempfile.gettempdir(),
        filename
    )


    doc.save(
        output
    )


    return output




@app.route("/")
def index():

    return render_template(
        "index.html"
    )



@app.route(
    "/generate",
    methods=["POST"]
)
def generate():


    # 客户可为空

    customer = request.form.get(
        "customer",
        ""
    )


    # 获取菜品文本

    dish_text = request.form.get(
        "dishes",
        ""
    )


    dishes = [
        x.strip()
        for x in dish_text.split("\n")
        if x.strip()
    ]


    file_path = create_word(
        customer,
        dishes
    )


    return send_file(
        file_path,
        as_attachment=True
    )




if __name__ == "__main__":


    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )


    app.run(
        host="0.0.0.0",
        port=port
    )
