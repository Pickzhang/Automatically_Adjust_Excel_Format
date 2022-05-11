# -*- coding = utf-8 -*-
# @Time : 2022/5/10 20:26
# @Author : Pickzhang
# @File : main.py
# @Software : PyCharm
from pathlib import Path
import xlwings as xw
from xlwings.utils import rgb_to_int

App = xw.App(visible=False, add_book=False)  # 创建APP
Current_path = Path.cwd()  # 获取当前路径
Folder_path = Path(Current_path)  # 将当前路径转换为可使用的路径格式
File_List = Folder_path.glob('*.xls*')  # 获取所有结尾是xls的文件名称列表

for Workbook_file in File_List:  # 遍历文件
    Workbook = App.books.open(Workbook_file)  # 打开工作簿

    for Worksheet in Workbook.sheets:  # 遍历工作表
        Worksheet_Used_Range = Worksheet.used_range  # 选中有数据的范围
        Worksheet_Used_Range.font.name = '微软雅黑'  # 设置字体名称
        Worksheet_Used_Range.font.size = 11  # 设置字体大小
        Worksheet_Used_Range.row_height = 18  # 设置行高
        Worksheet_Used_Range.rows.autofit()  # 行高自动调整

        for i in range(Worksheet_Used_Range.rows.count):  # 遍历选中范围的行
            if i == 0:  # 设置首行边框
                Worksheet_Used_Range.rows[i].api.Borders(8).LineStyle = 1
                Worksheet_Used_Range.rows[i].api.Borders(8).Weight = 2
                Worksheet_Used_Range.rows[i].api.Borders(8).Color = rgb_to_int((0, 0, 0))

            if i == Worksheet_Used_Range.rows.count - 1:  # 设置末行边框
                Worksheet_Used_Range.rows[i].api.Borders(9).LineStyle = 1
                Worksheet_Used_Range.rows[i].api.Borders(9).Weight = 2
                Worksheet_Used_Range.rows[i].api.Borders(9).Color = rgb_to_int((0, 0, 0))

            if i % 2 == 0:  # 隔行操作
                Worksheet_Used_Range.rows[i].color = (221, 235, 247)  # 添加底色
            else:
                Worksheet_Used_Range.rows[i].color = (255, 255, 255)

    Workbook.save()
    Workbook.close()
App.quit()
