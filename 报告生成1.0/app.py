from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt
from datetime import datetime
import pandas as pd

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate_report():
    # 获取表单数据
    name = request.form['name']
    date = request.form['date']
    data = request.form['data']

    # 将字符串数据转换为 JSON 格式
    import json
    data = json.loads(data)

    # 处理数据
    df = pd.DataFrame(data)

    # 创建一个新的 Word 文档
    doc = Document()

    # 添加标题
    title = doc.add_heading('报告', level=1)
    title.alignment = 1  # 居中对齐

    # 添加作者和日期
    doc.add_paragraph(f'作者: {name}')
    doc.add_paragraph(f'日期: {date}')

    # 添加表格
    table = doc.add_table(rows=1, cols=len(df.columns))
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = column

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value)

    # 保存文档
    filename = f'report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx'
    doc.save(filename)

    # 返回生成的文件
    return send_file(filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
