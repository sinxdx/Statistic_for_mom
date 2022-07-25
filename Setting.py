import pandas as pd
import os


class Setting:
    """定义设置类"""

    def __init__(self, path):
        route = os.path.join(path.main_path, "设置.xlsx")
        setting = pd.read_excel(route, sheet_name="设置", engine="openpyxl")
        setting.set_index("设置", inplace=True)
        self.S0 = setting.loc["S0", "选择"]
        self.S1 = setting.loc["S1", "选择"]
        self.S2 = setting.loc["S2", "选择"]
        self.S3 = setting.loc["S3", "选择"]
        self.S4 = setting.loc["S4", "选择"]
        self.S5 = setting.loc["S5", "选择"]
        self.S6 = setting.loc["S6", "选择"]  # keyword_enable
        print("已读取设置")
