import pandas as pd
from itertools import combinations

# 输入文件名称
infos_file_name = "信息.xlsx"
requirement_file_name = "证书要求.xlsx"
limit_file_name = "人员证书上限.xlsx"

# 加载输入
infos = pd.read_excel(infos_file_name,index_col=0)
requirements = pd.read_excel(requirement_file_name,index_col=0)
limits = pd.read_excel(limit_file_name,index_col=0)

# 校验输入
# 信息中的证书和证书要求一致
for one in infos.columns.values:
    if one not in requirements.columns.values:
        print("证书要求有问题，缺少 " + str(one) + " 的证书要求")
        exit
for one in infos.index.values:
    if one not in limits.index.values:
        print("人员证书上限有问题，缺少 " + str(one) + " 的证书上限设置")
        exit

all_possible = {}
# 遍历每一个人
for name in infos.index.values:
    limit = limits.loc[name, "证书上限"]
    # 遍历每个人可能的证书组合
    all = []
    for one in infos.loc[name, :]:
        if one != "" and pd.isna(one) is False:
            all.append(one)
    one_upper = min(limit, len(all))
    all_comb = combinations(all, one_upper)
    all_comb_ = []
    for one in all_comb:
        all_comb_.append(one)
    all_possible[name] = all_comb_

for one in all_possible:
    print(one)
    for two in all_possible[one]:
        print(two)

deltas = []

def generate_one_possible(all_possible, cur, name_list, i, requirements):
    name = name_list[i]
    for one in all_possible[name]:
        test = cur.copy()
        for cert in one:
            test.loc[name, cert] = cert
        if i == len(name_list) - 1:
            delta = 0
            # 遍历每一列，求每一列的和是否满足要求
            for cert in requirements.columns.values:
                count = 0
                for name in name_list:
                    if pd.isna(test.loc[name, cert]) is False:
                        count += 1
                if count < requirements.loc['目标', cert]:
                    print(cert + "不满足要求")
                    delta += (requirements.loc['目标', cert] - count)

            deltas.append([delta, test])
        else:
            generate_one_possible(all_possible, test, name_list, i+1, requirements)

def compare_(x, y):
    return x[0] - y[0]

def dump_to_xlsx(cur, file_name):
    # 增加两行，一行是目标，一行是现状
    # 增加两列，一列是目标，一列是现状
    counts = []
    for cert in infos.columns.values:
        count = 0
        for name in infos.index.values:
            if pd.isna(cur.loc[name, cert]) is False:
                count += 1
        counts.append(count)
    cur.loc["现状"] = counts
    cur.loc["目标"] = requirements.loc['目标',:]

    counts = []
    for name in infos.index.values:
        count = 0
        for cert in infos.columns.values:
            if pd.isna(cur.loc[name, cert]) is False:
                count += 1
        counts.append(count)
    counts.append("")
    counts.append("")
    cur.loc[:, "现状"] = counts
    temp = limits.loc[:, "证书上限"].to_list()
    temp.append("")
    temp.append("")
    cur.loc[:, "目标"] = temp

    cur.to_excel(file_name + ".xlsx")


generate_one_possible(all_possible, pd.DataFrame(data=None, columns=infos.columns.values, index=infos.index.values), infos.index.values.tolist(), 0, requirements)
sorted_deltas = sorted(deltas, key=(lambda x:x[0]))
for i in range(min(3, len(sorted_deltas))):
    dump_to_xlsx(sorted_deltas[i][1], str(i))
