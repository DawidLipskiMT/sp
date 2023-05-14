import numpy as np
import pandas as pd


def parser(t: pd.DataFrame, x) -> pd.DataFrame:
    if 'ID' not in t.columns:
        t = pd.DataFrame({'ID': 1,
                          'Nazwa': t.iloc[:, 0] + ' ' + t.iloc[:, 1],
                          'Perfumeria': t.iloc[:, 2]})
    if 'test' in t.columns:
        t.dropna(subset=['test'], inplace=True)
    t.iloc[:, 2] = t.iloc[:, 2].apply(lambda perf_num: extract_perf_number(perf_num))
    t.drop_duplicates(subset=['Nazwa', 'Perfumeria'], inplace=True)
    x = prepare_matrix_with_places(x)
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


def prepare_matrix_with_places(x: pd.DataFrame):
    x = x.iloc[:, :].reset_index(drop=True)
    #x.columns = x.iloc[0, :]
    x = x.iloc[:, :].reset_index(drop=True)
    return x


def fill_columns_with_names(x, t, place):
    temp = t[t.iloc[:, 2] == place]
    values = temp.iloc[:, 1]
    perf_index = x.index[x.iloc[:, 0] == place]
    for i in range(len(values)):
        x.iloc[perf_index[0], 3 + i] = values.iloc[i]
    return x


if __name__ == '__main__':
    v0 = pd.read_excel('SPH 22-19-10.xls')
    v1 = pd.read_excel('s_dystr_2021.xls')
    df = parser(v0, v1)
    df.to_excel("SPH 22-19-10-fin.xlsx", index=False)
