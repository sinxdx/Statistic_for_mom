import pandas as pd

from Path import Path
from Setting import Setting
from fuction import 统计程序, load_统计字典, load_关键词字典, load_分隔符列表, raw_input, traverse_xlsx
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
        keyword = input("输入你想读取的关键词子表(MR/CT/DR/...)")
        '''根据统计的类别，读取关键词词典，然后读取对应的关键词和分隔符'''
        dc_raw = load_关键词字典(path, keyword)
        weight = load_统计字典(dc_raw)
        split_list = load_分隔符列表(dc_raw)
        file_path_list = []
        cnt_all = 0
        '''读入需要统计的数据'''
        if setting.S1 == "是":  # 如果设置了需要搜索多个文件，那么就进入文件夹里搜索
            try:
                file_path_list = traverse_xlsx(route=path.data_path, setting=setting)
            finally:
                os.chdir(path.main_path)
        else:
            file_path_list = [raw_input("把你导出的数据拖进来，注意文件存放的路径中，文件夹名字不能以数字开头")]
        while True:
            try:
                for file_path in file_path_list:
                    print(f"统计文件{file_path}")
                    cnt = 统计程序(file_path=file_path, keywords=keyword, weight=weight, split_list=split_list,
                               setting=setting, path=path)
                    cnt_all += cnt
                print(f"以上全部的统计结果为{cnt_all}")
                break
            except FileNotFoundError:
                file_path_list = [raw_input("文件读取未成功，请再次把文件拖进来")]
                continue
