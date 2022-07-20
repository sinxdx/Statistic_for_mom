import pandas as pd
import os


def 统计程序(data_raw, keywords, weight, split_list, setting):
    mode = setting.S0
    if mode == "新院区":
        data = data_raw.loc[data_raw["检查类型"] == keywords, "检查全部项目"]
    elif mode == "旧院区":
        data = data_raw.loc[data_raw["设备"] == keywords, "检查部位"]

    data.dropna(how="any", inplace=True)
    data.reset_index(drop=True, inplace=True)

    not_statistic_items = []  # 未被统计到的单元格
    not_statistic_words = []  # 未被统计到的字段
    cnt = 0
    for i in data.loc[:]:
        temp = False
        # print(f"——————检查项目{i}————————")
        ls = split(i, split_list)
        for ts in ls:  # ts == "头颅CT平扫+增强"
            temp2 = False
            for (key, values) in weight.items():
                if key in ts:
                    temp = True
                    temp2 = True
                    cnt += values
                    if setting.S2 == "是":
                        print(f"统计到{key},其值为{values}")
                    break
            if not temp2 and setting.S3 == "是":
                not_statistic_words.append(ts)
        if not temp and setting.S4 == "是":
            not_statistic_items.append(i)
    if setting.S3 == "是":
        t = set(not_statistic_words)
        print(f'{t}未被统计到')
        del t
    if setting.S4 == "是":
        t = set(not_statistic_items)
        print(f'{t}未被统计到')
        del t
    print(f"文件统计总和为{cnt}")
    return cnt


"""def  load_导出数据(path):
    file = pd.read_excel(file_path)
    return file """


def load_关键词字典(path, keyword):
    while not keyword:
        keyword = input("输入你需要统计的类别(MR/CT/DR)")
    try:
        route = os.path.join(path.main_path, "关键词字典.xlsx")
        dc_file = pd.ExcelFile(route,engine="openpyxl")
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
