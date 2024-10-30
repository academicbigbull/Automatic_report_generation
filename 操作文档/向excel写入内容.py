#修改内容直接赋值就行
from openpyxl import load_workbook
workbook = load_workbook(filename = "E:/python/操作文档/test.xlsx")
sheet=workbook['Sheet1'] #获取sheet1表
sheet=workbook.active #打开获取的表
cell1=sheet['A1']
cell1.value='shit'
workbook.save(filename = "E:/python/操作文档/test.xlsx")

#向表中插入行数据
data = [
    ["唐僧","男","180cm"],
    ["孙悟空","男","188cm"],
    ["猪八戒","男","175cm"],
    ["沙僧","男","176cm"],
]
for row in data:
    sheet.append(row)
workbook.save(filename='E:/python/操作文档/test.xlsx')