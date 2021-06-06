from xlsxwriter.utility import xl_range
import pandas as pd
import pandas.io.formats.excel
import xlsxwriter


def decoratorToExcel(excel_name: str = "test", excel_title: str = "test", excel_dir: str = None):
    """
    输出文件装饰器,作用于类方法
    :param excel_dir: 输出的文件夹
    :param excel_title: 文件名
    :param excel_name: 文件表头
    :return:
    """

    def midFunc(cls_func):

        def inner(self, *args, **kwargs):
            data: pd.DataFrame = cls_func(self, *args, **kwargs)
            writer = pd.ExcelWriter(
                '%s%s.xlsx' % (
                    excel_dir + "/" if excel_dir else "", excel_name,), datetime_format='yyyy/mm/dd')
            data.to_excel(
                writer,
                engine='xlsxwriter',
                sheet_name='sheet',
                startrow=2,
                header=False,
                index=False,
                float_format="%.2f",
            )
            workbook = writer.book
            worksheet = writer.sheets['sheet']
            worksheet.set_row(0, 20)  # 行高

            header_format = workbook.add_format({
                'bold': True,
                'font_color': 'black',
                'text_wrap': True,
                'align': 'left',
                'fg_color': '#6BA81E',
                'border': 1})
            rl = data.columns.tolist()
            print(rl)
            for col_num, value in enumerate(rl):
                print(col_num)
                print(value)
                worksheet.write(0,col_num, value, header_format)
            writer.save()
            return

        return inner

    return midFunc


class Test:

    @decoratorToExcel(excel_name="测试表格", excel_title="测试sheet")
    def setExcelData(self) -> pd.DataFrame:
        """
        生成excel数据
        :return: DataFrame
        """
        data: pd.DataFrame = pd.DataFrame({"name": "张三 李四 王五".split(), "age": "12 13 15".split()})
        return data


if __name__ == '__main__':
    Test().setExcelData()
