import os


class Path:
    """定义路径类"""

    def __init__(self):
        self.main_path = os.getcwd()
        self.data_path = os.path.join(self.main_path, "导出数据存储处")
        self.report_path = os.path.join(self.main_path, "统计结果保存处")

    def make_dir(self):
        if not os.path.exists(self.data_path):
            os.makedirs(self.data_path)
