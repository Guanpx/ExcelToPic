import xlwings as xw
# from xlwings.constants import CopyPictureFormat, PictureAppearance

from PIL import ImageGrab

# https://github.com/xlwings/xlwings/issues/582

def excel_save_img(path, img_path, sheet=0, img_name="1", img_suffix="png"):
    app = xw.App(visible=False, add_book=False)
    # 1. 使用 xlwings 的 读取 path 文件 启动
    wb = app.books.open(path)

    # 2. 读取 sheet
    sht = wb.sheets[sheet]

    # 3. 获取 行与列
    nrow = sht.api.UsedRange.Rows.count
    ncol = sht.api.UsedRange.Columns.count
    print(nrow)
    print(ncol)

    # 4. 获取有内容的 range
    range_val = sht.range(
        (1, 1),  # 获取 第一行 第一列
        (nrow, ncol)  # 获取 第 nrow 行 第 ncol 列
    )
    print(range_val.value)

    # 5. 复制图片区域
    range_val.api.CopyPicture()

    # 6. 粘贴
    sht.api.Paste()

    pic = sht.pictures[0]  # 当前图片
    pic.api.Copy()  # 复制图片

    img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
    img.save(img_name + "." + img_suffix)  # 保存图片
    pic.delete()  # 删除sheet上的图片

    wb.close()  # 不保存，直接关闭
    app.quit()  # 退出


if __name__ == '__main__':
    p = r"C:\Users\Magic\PycharmProjects\pythonProject\pandas_out.xlsx"
    excel_save_img(p, r"C:\Users\Magic\PycharmProjects\pythonProject")
