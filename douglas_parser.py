import numpy as np
import pandas as pd


def parser(t, x) -> pd.DataFrame:
    t.iloc[:, 2] = t.iloc[:, 2].apply(lambda perf_num: extract_perf_number(perf_num))
    n_col = number_of_columns_to_add(t)
    x = add_columns(n_col, x)
    places = x.iloc[:, 0].unique()
    for place in places:
        x = fill_columns_with_names(x, t, place)
    return x


def extract_perf_number(perf_num):
    numbers = []
    if isinstance(perf_num, int):
        return perf_num
    if isinstance(perf_num, float):
        return int(perf_num)
    for word in perf_num.split():
        if word.isdigit():
            numbers.append(int(word))
    return numbers[0]


def number_of_columns_to_add(report):
    return report.groupby(by='Filia').count()['Nazwisko'].values.max()


def add_columns(n_col, t):
    d = t.copy()
    for i in range(n_col):
        d[f'Imię_Nazwisko_{i + 1}'] = ''
    return d


def fill_columns_with_names(x, t, place):
    temp = t[t.iloc[:, 2] == place]
    values = temp.iloc[:, 0] + ' ' + temp.iloc[:, 1]
    perf_index = x.index[x.iloc[:, 0] == place]
    for i in range(len(values)):
        x.iloc[perf_index[0], 15 + i] = values.iloc[i]
    return x


if __name__ == '__main__':
    v0 = pd.read_excel('DGL 22-12-10.xls')
    v1 = pd.read_excel('d_pos_07_2020.xls')
    df = parser(v0, v1)
    df.to_excel("DGL 22-12-10-fin.xlsx", index=False)
