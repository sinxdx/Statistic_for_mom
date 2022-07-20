import pandas as pd

from Path import Path
from Setting import Setting
from fuction import 统计程序, load_统计字典, load_关键词字典, load_分隔符列表, raw_input
import os
path = Path()
path.make_dir()
setting = Setting(path)

while True:
    k = input("直接回车，开始运行程序\n"
              "输入1并回车，读取设置\n"
              "输入q并回车，退出程序\n")
    if k == "1":
        setting = Setting(path=path)
    elif k == "q":
        exit()
    else:
        '''首先决定选择统计哪个类别'''
        keyword = input("输入你需要统计的类别(MR/CT/DR)")
        '''根据统计的类别，读取关键词词典，然后读取对应的关键词和分隔符'''
        dc_raw = load_关键词字典(path, keyword)
        weight = load_统计字典(dc_raw)
        split_list = load_分隔符列表(dc_raw)
        file_path_list = []
        '''读入需要统计的数据'''
        if setting.S1 == "是":
            try:
                os.chdir(path.data_path)
                for folder in os.listdir():
                    if keyword in folder:
                        os.chdir(folder)
                        for file in os.listdir():
                            if str(setting.S5) in file:  # 添加统计年份关键词，用于快速统计年份
                                file_path_list.append(file)
                            else:
                                print(f"文件<{file}>中不包含“{setting.S5}”，已跳过")
            finally:
                os.chdir(path.main_path)
        else:
            file_path_list = [raw_input("把你导出的数据拖进来，注意文件存放的路径中，文件夹名字不能以数字开头")]
        while True:
            try:
                for file_path in file_path_list:
                    print(f"读取文件{file_path}")
                    file = pd.ExcelFile(file_path)
                    data_frame = pd.read_excel(file)
                    统计程序(data_raw=data_frame, keywords=keyword, weight=weight, split_list=split_list, setting=setting)
                    k = input("统计完成，按下回车以继续")
                    break
            except FileNotFoundError:
                file_path = raw_input("文件读取未成功，请再次把文件拖进来")
                continue
            except KeyError:
                print('查看这份表的表头是不是缺了“检查类型"和"检查全部项目"中的某一项\n如果是旧版的，看看是不是缺了"设备"或者"检查部位"')
                file_path = raw_input("请重新把正确的文件拖进来")
                continue
