2024-12-17 关于获取指定特殊报告获取它的特征关键
from docx import Document
from jinja2 import Template


def extract_docx_features(file_path):
    # 读取 docx 文件
    doc = Document("E://python//获取报告生成适配该报告的jinja2模板//rest_1.docx")

    # 初始化特征字典
    features = {
        "title": None,
        "title_color": None,  # 标题颜色
        "title_size": None,  # 标题字体大小
        "report_people": None,  # 汇报人
        "report_people_color": None,  # 汇报人字体颜色
        "report_people_size": None,  # 汇报人字体大小
        "report_datatime": None,  # 实验日期
        "report_datatime_color": None,  # 实验日期字体颜色
        "report_datatime_size": None,  # 实验日期字体大小
        "report_place": None,  # 实验地点
        "report_place_color": None,  # 实验地点字体颜色
        "report_place_size": None,  # 实验地点字体大小
        # 表格和图片不用设置格式
        "data_chart": [],
        "data_picture": None
    }

    # 提取标题信息
    if doc.paragraphs:
        title_paragraph = doc.paragraphs[0]
        features["title"] = title_paragraph.text
        features["title_color"] = title_paragraph.style.font.color.rgb if title_paragraph.style.font.color else None
        features["title_size"] = title_paragraph.style.font.size.pt if title_paragraph.style.font.size else None

    # 提取汇报人、实验日期、实验地点等信息
    for para in doc.paragraphs:
        if '汇报人' in para.text:
            features["report_people"] = para.text.split('：')[1].strip()  # 假设格式是 "汇报人：XXX"
            features["report_people_color"] = para.style.font.color.rgb if para.style.font.color else None
            features["report_people_size"] = para.style.font.size.pt if para.style.font.size else None
        elif '实验日期' in para.text:
            features["report_datatime"] = para.text.split('：')[1].strip()  # 假设格式是 "实验日期：YYYY-MM-DD"
            features["report_datatime_color"] = para.style.font.color.rgb if para.style.font.color else None
            features["report_datatime_size"] = para.style.font.size.pt if para.style.font.size else None
        elif '实验地点' in para.text:
            features["report_place"] = para.text.split('：')[1].strip()  # 假设格式是 "实验地点：XXX"
            features["report_place_color"] = para.style.font.color.rgb if para.style.font.color else None
            features["report_place_size"] = para.style.font.size.pt if para.style.font.size else None

    # 提取表格内容
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            table_data.append(row_data)
        features["data_chart"].append(table_data)

    # 提取图片信息（获取图片的路径或URL）
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            features["data_picture"] = rel.target_ref  # 提取图片路径或URL

    return features


def generate_jinja2_template(features):
    # 生成动态的 Jinja2 模板
    template = """
    <!DOCTYPE html>
    <html lang="zh">
    <head>
        <meta charset="UTF-8">
        <title>{{ title }}</title>
        <style>
            body {
                font-family: "Arial", sans-serif;
            }
            h1 {
                font-size: {{ title_size }}pt;
                color: {{ title_color }};
            }
            .report-people {
                font-size: {{ report_people_size }}pt;
                color: {{ report_people_color }};
            }
            .report-datetime {
                font-size: {{ report_datatime_size }}pt;
                color: {{ report_datatime_color }};
            }
            .report-place {
                font-size: {{ report_place_size }}pt;
                color: {{ report_place_color }};
            }
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th, td {
                border: 1px solid black;
                padding: 8px;
                text-align: center;
            }
        </style>
    </head>
    <body>
        {% if title %}
        <h1>{{ title }}</h1>
        {% endif %}

        {% if report_people %}
        <p class="report-people">汇报人: {{ report_people }}</p>
        {% endif %}

        {% if report_datatime %}
        <p class="report-datetime">实验日期: {{ report_datatime }}</p>
        {% endif %}

        {% if report_place %}
        <p class="report-place">实验地点: {{ report_place }}</p>
        {% endif %}

        {% if data_chart %}
        <h2>实验数据</h2>
        {% for table in data_chart %}
            <table>
                <thead>
                    <tr>
                        {% for column in table[0] %}
                            <th>{{ column }}</th>
                        {% endfor %}
                    </tr>
                </thead>
                <tbody>
                    {% for row in table[1:] %}
                        <tr>
                            {% for cell in row %}
                                <td>{{ cell }}</td>
                            {% endfor %}
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        {% endfor %}
        {% endif %}

        {% if data_picture %}
        <h2>报告图片</h2>
        <img src="{{ data_picture }}" alt="报告图片" />
        {% endif %}

    </body>
    </html>
    """

    return template


# 使用示例
file_path = "E://python//获取报告生成适配该报告的jinja2模板//rest_1.docx"
report_features = extract_docx_features(file_path)

# 生成 Jinja2 模板
template = generate_jinja2_template(report_features)

# 使用 Jinja2 渲染模板
jinja_template = Template(template)
rendered_html = jinja_template.render(report_features)

# 打印渲染后的 HTML
print(rendered_html)


报告内容为：
空气动力学实验报告
汇报人：丁真
实验日期：2024-12-17
实验地点：西南科技大学
数据表格：
ID	Name	Age	Department	Salary
1	Alice	30	Engineering	70000
2	Bob	24	Marketing	50000
3	Charlie	28	Human Resources	55000
4	David	35	Finance	80000
5	Eve	32	Engineering	75000
数据图：xxx


输出的更改后的模板内容为：
    <!DOCTYPE html>
    <html lang="zh">
    <head>
        <meta charset="UTF-8">
        <title>空气动力学实验报告</title>
        <style>
            body {
                font-family: "Arial", sans-serif;
            }
            h1 {
                font-size: 12.0pt;
                color: None;
            }
            .report-people {
                font-size: 12.0pt;
                color: None;
            }
            .report-datetime {
                font-size: 12.0pt;
                color: None;
            }
            .report-place {
                font-size: 12.0pt;
                color: None;
            }
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th, td {
                border: 1px solid black;
                padding: 8px;
                text-align: center;
            }
        </style>
    </head>
    <body>
        
        <h1>空气动力学实验报告</h1>
        

        
        <p class="report-people">汇报人: 丁真</p>
        

        
        <p class="report-datetime">实验日期: 2024-12-17</p>
        

        
        <p class="report-place">实验地点: 西南科技大学</p>
        

        
        <h2>实验数据</h2>
        
            <table>
                <thead>
                    <tr>
                        
                            <th>ID</th>
                        
                            <th>Name</th>
                        
                            <th>Age</th>
                        
                            <th>Department</th>
                        
                            <th>Salary</th>
                        
                    </tr>
                </thead>
                <tbody>
                    
                        <tr>
                            
                                <td>1</td>
                            
                                <td>Alice</td>
                            
                                <td>30</td>
                            
                                <td>Engineering</td>
                            
                                <td>70000</td>
                            
                        </tr>
                    
                        <tr>
                            
                                <td>2</td>
                            
                                <td>Bob</td>
                            
                                <td>24</td>
                            
                                <td>Marketing</td>
                            
                                <td>50000</td>
                            
                        </tr>
                    
                        <tr>
                            
                                <td>3</td>
                            
                                <td>Charlie</td>
                            
                                <td>28</td>
                            
                                <td>Human Resources</td>
                            
                                <td>55000</td>
                            
                        </tr>
                    
                        <tr>
                            
                                <td>4</td>
                            
                                <td>David</td>
                            
                                <td>35</td>
                            
                                <td>Finance</td>
                            
                                <td>80000</td>
                            
                        </tr>
                    
                        <tr>
                            
                                <td>5</td>
                            
                                <td>Eve</td>
                            
                                <td>32</td>
                            
                                <td>Engineering</td>
                            
                                <td>75000</td>
                            
                        </tr>
                    
                </tbody>
            </table>
        
        

        
        <h2>报告图片</h2>
        <img src="media/image1.png" alt="报告图片" />
        

    </body>
    </html>
