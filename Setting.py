import pandas as pd
import os


class Setting:
    """定义设置类"""

    def __init__(self, path):
        route = os.path.join(path.main_path, "设置.xlsx")
        setting = pd.read_excel(route, sheet_name="设置")
        setting.set_index("设置", inplace=True)
        self.导出数据存储处 = setting.loc["直接在文件夹<导出数据存储处>中依次读取文件", "选择"]
        self.每统计一行都输出一次 = setting.loc["每统计一行都输出一次", "选择"]
        self.显示没被统计到的字段 = setting.loc["显示没被统计到的字段", "选择"]
        self.显示没被统计到的单元格 = setting.loc["显示没被统计到的单元格", "选择"]

        print("已读取设置")
