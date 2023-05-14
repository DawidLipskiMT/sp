import numpy as np
import pandas as pd

def extract_perf_number(perf_num):
    numbers = []
    if isinstance(perf_num, int):
        return perf_num
    for word in perf_num.split():
        if word.isdigit():
            numbers.append(int(word))
    return numbers[0]


def prepare_matrix_with_places(x):
    x = x.iloc[1:, :]
    x.columns = x.iloc[1]
    x.drop(1, inplace=True)
    return x


def fill_columns_with_names(x, t, place):
    temp = t[t.iloc[:, 2] == place]
    values = temp.iloc[:, 1]
    perf_index = x.index[x.iloc[:, 0] == place]
    for i in range(len(values)):
        x.iloc[perf_index[0], 3 + i] = values.iloc[i]
    return x


def test_parser(t, x) -> pd.DataFrame:
    t.iloc[:, 2] = t.iloc[:, 2].apply(lambda perf_num: extract_perf_number(perf_num))
    x = prepare_matrix_with_places(x)
    places = x.iloc[:, 0].unique()
    for place in places:
        x = fill_columns_with_names(x, t, place)
    return x


if __name__ == '__main__':
    v0 = pd.read_excel('t.xlsx')
    v1 = pd.read_excel('x.xlsx')
    df = test_parser(v0, v1)
    df.to_csv("output.xlsx")