import pandas as pd
import os

from xlrd import XLRDError


def 统计程序(file_path, keywords, weight, split_list, setting, path):
    """对文件进行统计并返回统计值，同时还有生成报告的功能"""
    '''读取文件,关键在于.xlsx和.xls不能使用同一个engine'''
    try:
        data_raw = pd.read_excel(file_path, engine="xlrd")
    except XLRDError:
        data_raw = pd.read_excel(file_path, engine="openpyxl")
    except PermissionError:
        print(f"访问被拒绝\n请关闭当前打开的文件{os.path.basename(file_path)}")
        return 0
    mode = setting.S0  # 新院区与旧院区导出程序的区别
    keyword_enable = setting.S6
    try:
        if keyword_enable == "使用关键词子表(MR/CT/DR/...)":
            if mode == "旧院区":
                data = data_raw.loc[data_raw["设备"] == keywords, ["检查部位", "姓  名"]]
            elif mode == "新院区":
                data = data_raw.loc[data_raw["检查类型"] == keywords, ["检查全部项目", "患者姓名"]]
            elif mode == "新院区2":
                data = data_raw.loc[data_raw["检查类型"] == keywords, ["检查项目", "患者姓名"]]
        elif keyword_enable == "自行输入":
            temp_keyword = input("输入你想要筛选出的”检查类型“的关键字")
            if mode == "旧院区":
                data = data_raw.loc[data_raw["设备"] == temp_keyword, ["检查部位", "姓  名"]]
            elif mode == "新院区":
                data = data_raw.loc[data_raw["检查类型"] == temp_keyword, ["检查全部项目", "患者姓名"]]
            elif mode == "新院区2":
                data = data_raw.loc[data_raw["检查类型"] == temp_keyword, ["检查项目", "患者姓名"]]
        elif keyword_enable == "不进行筛选":
            if mode == "旧院区":
                data = data_raw.loc[:, ["检查部位", "姓  名"]]
            elif mode == "新院区":
                data = data_raw.loc[:, ["检查全部项目", "患者姓名"]]
            elif mode == "新院区2":
                data = data_raw.loc[:, ["检查项目", "患者姓名"]]
    except KeyError:
        print("关键词表头读取错误,检查你的文件是否正确\n"
              f"或到《设置.xlsx》中检查一下自己的关键词表头是否选择正确\n"
              f"你当前选择的设置项是{mode}，请到《设置.xlsx》中检查一下自己的关键词表头是否选择正确\n")
        return 0

    data.dropna(how="any", inplace=True)
    data.reset_index(drop=True, inplace=True)

    report = pd.DataFrame(columns=["检查全部项目", "该单元格统计出的部位数", "该单元格的字段", "字段统计明细"])

    not_statistic_items = []  # 显示未统计到的单元格
    not_statistic_words = []  # 显示未统计到的字段
    cnt_all = 0  # 用来输出统计值
    loc = 0  # 用作指向"检查全部项目"的指针
    for item in data.loc[:, ""]:
        report.loc[loc, "检查全部项目"] = item  # 报告中添加"检查全部项目"一项，其位置指针应当指向loc

        cnt_in_item = 0  # 单元格内的统计部位数
        loc_in_item = 0  # 字段在单元格中的位置，理论上应从1开始，使得统计时能与检查全部项目错开一格,但下面把+=1放在for循环的前面了所以还是从0开始
        item_output_flag = False  # 单元格是否被统计到的标志位
        ls = split(item, split_list)
        for word in ls:  # ts == "头颅CT平扫+增强"
            loc_in_item += 1
            report.loc[loc + loc_in_item, "该单元格的字段"] = word  # 在报告中添加"该单元格的字段"一项
            word_output_flag = False  # 字段是否被统计到的标志位
            for (key, values) in weight.items():
                if key in word:
                    item_output_flag = True
                    word_output_flag = True
                    cnt_in_item += values
                    report.loc[loc + loc_in_item, "字段统计明细"] = f"统计到{key},其值为{values}"
                    break
            if not word_output_flag and setting.S3 == "是":
                report.loc[loc + loc_in_item, "字段统计明细"] = "未统计到关键词"
                not_statistic_words.append(word)
        if not item_output_flag and setting.S4 == "是":
            report.loc[loc, "该字段统计明细"] = "该单元格内未统计到关键词"
            not_statistic_items.append(item)
        report.loc[loc, "该单元格统计出的部位数"] = cnt_in_item  # 在每个item统计结束后，输出在item中统计出的单元格数量
        cnt_all += cnt_in_item  # 同时文件内的统计总数也要加上
        '''在每个item内部的for循环结束后，需要更新loc指针在report中的位置'''
        loc += loc_in_item  # 更新loc指针的指向
        loc += 1
    if setting.S3 == "是":
        t = set(not_statistic_words)
        print(f'{t}未被统计到')
        del t
    if setting.S4 == "是":
        t = set(not_statistic_items)
        print(f'{t}未被统计到')
        del t
    if setting.S2 == "是":
        try:
            report.to_excel(f"{os.path.join(path.report_path, os.path.basename(file_path))}_{keywords}统计明细.xlsx")
        except PermissionError:
            print(f"访问被拒绝\n请关闭文件<{os.path.basename(file_path)}_{keywords}统计明细.xlsx>")
    print(f"《{file_path}》的统计结果为{cnt_all}")
    return cnt_all


"""def  load_导出数据(path):
    file = pd.read_excel(file_path)
    return file """


def load_关键词字典(path, keyword):
    while not keyword:
        keyword = input("输入你需要统计的类别(MR/CT/DR)")
    try:
        route = os.path.join(path.main_path, "关键词字典.xlsx")
        dc_file = pd.ExcelFile(route, engine="openpyxl")
    except FileNotFoundError:
        route = raw_input("在当前文件中未找到关键词字典.xlsx，请把你的关键词字典.xlsx拖到这里来")
        dc_file = pd.ExcelFile(route)
    while not (keyword in dc_file.sheet_names):
        keyword = input(f"该类别{keyword}的关键词未在关键词字典中定义，请重新输入关键词，或到请到关键词词典.xlsx中对该关键词进行定义\n"
                        f"请重新重新输入类别名（CT/MR/DR...)")
    dc_raw = pd.read_excel(dc_file, sheet_name=keyword)
    return dc_raw


def load_统计字典(dc_raw):
    dc = dc_raw.loc[:, ["要统计的关键词", "统计到关键词时计算的次数"]]
    dc = dc.dropna(how="all")
    dc = dc.set_index("要统计的关键词")
    dc = dc.to_dict(orient="series")
    dc = dc["统计到关键词时计算的次数"].to_dict()
    return dc


def load_分隔符列表(dc_raw):
    split_list = dc_raw.loc[:, "分隔符"].dropna().to_list()
    return split_list


def split(lss, spl):
    for i in spl:
        lss = lss.replace(i, "$")
    return lss.split("$")


def raw_input(message):
    st = input(message)
    st = st.replace("\"", "")
    st = st.replace("\\", "/")
    return st


def traverse_xlsx(route, setting):
    """用于遍历文件夹里的所有.xlsx/.xls文件"""
    ls = []  # 将返回的路径存储在这个ls中

    try:
        os.chdir(route)
    except NotADirectoryError:  # 若该文件不是文件夹,那么也没有必要再往下查找了，直接返回空值
        return ls

    '''if用来过滤年份'''
    if setting.S5 == "全部":
        flag = ""
    else:
        flag = str(setting.S5)
    for file in os.listdir():
        if file.endswith(".xlsx") and (flag in file) and not ("~$" in file):
            # print(f"处理文件{file}")
            ls.append(os.path.join(os.getcwd(), file))
        elif file.endswith(".xls") and (flag in file):
            # print(f"处理文件{file}")
            ls.append(os.path.join(os.getcwd(), file))
        else:  # 若file是个文件夹，那就进入文件夹搜索
            ls = ls + traverse_xlsx(file, setting)
    return ls
