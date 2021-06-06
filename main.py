import pandas as pd
import pandas.io.formats.excel
import xlwings as xw

from win32com.client import Dispatch, DispatchEx
import pythoncom
from PIL import ImageGrab, Image
import datetime

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def test():
    d = pd.DataFrame({'a': ['a', '1', 'b', '2'],
                      'b': ['a', 'b', 'c', 'd'],
                      'c': [1, 2, 3, 4]})

    writer = pd.ExcelWriter('./pandas_out.xlsx', engine='xlsxwriter')

    d.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=1)

    # worksheet = writer.sheets['Sheet1']
    #
    # format = workbook.add_format()
    # format2 = workbook.add_format()
    # format.set_align('right')
    # format2.set_align('left')
    # # worksheet.set_column('A:D', 11, format)
    # rs = worksheet.set_row(1, cell_format=format)

    # worksheet.set_header()
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # 3. 获取 行与列
    # nrow = worksheet.api()
    # ncol = worksheet
    # print(nrow)
    # print(ncol)
    #
    # # 4. 获取有内容的 range
    # range_val = worksheet.range(
    #     (1, 1),  # 获取 第一行 第一列
    #     (nrow, ncol)  # 获取 第 nrow 行 第 ncol 列
    # )
    # print(range_val.value)
    #
    # # 5. 复制图片区域
    # range_val.api.CopyPicture()

    border_fmt = workbook.add_format({'align': 'left',
                                      'valign': 'vcenter',
                                      'bold': False,
                                      'font_color': 'black',
                                      'font_size': '16',
                                      'text_wrap': True,
                                      'border': 0})
    hd = d.columns.tolist()
    for col_num, value in enumerate(hd):
        print(col_num)
        print(value)
        worksheet.write(0, col_num, value, border_fmt)

    worksheet.set_column('A:D', 11, border_fmt)
    writer.save()
    pass


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    test()
