from flask import Flask
from lxml import etree

app = Flask(__name__)

def clean_xml_data(xml_file_path):
    tree = etree.parse(xml_file_path)
    root = tree.getroot()

    title = root.find('title').text if root.find('title') is not None else ''
    author = root.find('author').text if root.find('author') is not None else ''
    date = root.find('date').text if root.find('date') is not None else ''

    sections = []
    for section in root.findall('sections/section'):
        heading = section.find('heading').text if section.find('heading') is not None else ''
        content = section.find('content').text if section.find('content') is not None else ''
        sections.append({'heading': heading, 'content': content})

    data_table = {}
    table = root.find('dataTable')
    if table is not None:
        table_name = table.find('tableName').text if table.find('tableName') is not None else ''
        columns = [col.text for col in table.findall('columns/column')]
        rows = []
        for row in table.findall('rows/row'):
            cells = [cell.text for cell in row.findall('cell')]
            rows.append(cells)
        data_table = {'tableName': table_name, 'columns': columns, 'rows': rows}

    return {
        'title': title,
        'author': author,
        'date': date,
        'sections': sections,
        'dataTable': data_table
    }

@app.route('/')
def index():
    xml_file_path = 'static/data/report_data.xml'  # 替换为实际的 XML 文件路径
    cleaned_data = clean_xml_data(xml_file_path)
    return f"Cleaned data: {cleaned_data}"

if __name__ == '__main__':
    app.run()