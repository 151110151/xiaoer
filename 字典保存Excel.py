# -*- coding: UTF-8 -*-
"""
@Project ：新闻 
@File ：字典保存Excel.py
@Author ：xiaoer
@Date ：2021/10/27 10:37 
@desc:
"""
# -*- coding: utf-8 -*-
import xlsxwriter


# 生成excel文件
def generate_excel(expenses):
    workbook = xlsxwriter.Workbook('经纬度.xlsx')
    worksheet = workbook.add_worksheet()

    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
    # bold：加粗，num_format:数字格式
    bold_format = workbook.add_format({'bold': True})
    # money_format = workbook.add_format({'num_format': '$#,##0'})
    # date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

    # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15)

    # 用符号标记位置，例如：A列1行
    worksheet.write('A1', '城市名', bold_format)
    worksheet.write('B1', '坐标', bold_format)
    row = 1
    col = 0
    for item in expenses:
        # 使用write_string方法，指定数据格式写入数据
        worksheet.write_string(row, col, str(item['以城市名检索']))
        worksheet.write_string(row, col + 1, item['坐标'])
        row += 1
    workbook.close()


if __name__ == '__main__':
    rec_data = [
        {'以城市名检索': "Alta Floresta D'Oeste", '坐标': '111,222'},
        {'以城市名检索': "Alta Floresta ", '坐标': '24,222'},
        {'以城市名检索': "Alta Floresta D", '坐标': '111,42'},
        {'以城市名检索': "Alta Floresta D'Oese", '坐标': '427,222'},
    ]
    generate_excel(rec_data)
