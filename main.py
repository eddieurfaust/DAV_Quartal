from typing import List

import pandas as pd
from path import Path


class Worker:
    def __init__(self, df_: pd.DataFrame):
        self.df = df_
        self._clean_dataframe()

    def _drop(self, cols: list, rows: list):
        self.df.drop(cols, axis=1, inplace=True)
        self.df.drop(rows, axis=0, inplace=True)

        return self

    def _set_column_headers(self):
        self.df.columns = self.df.iloc[0].to_list()

        self._drop(cols=[], rows=[self.df.index[0]])

        return self

    def _replace_NaN_with_None(self):
        self.df.where(pd.notnull(self.df), None, inplace=True)

        return self

    def _remove_rows_by_substring(self, lst: List[str]):
        for substr in lst:
            self.df = self.df[~self.df['Kurs'].str.contains(substr).fillna(value=False)]

        return self

    def _drop_na(self):
        self.df.dropna(how='all', axis=0, inplace=True)  # remove cols
        self.df.dropna(how='all', axis=1, inplace=True)  # remove rows

        return self

    def _clean_dataframe(self):
        self._drop(cols=['Unnamed: 0', 'None.7'], rows=[0, 1])  # cols würden auch durch self.drop_na() weg fallen?
        self._set_column_headers()
        self._replace_NaN_with_None()
        self._remove_rows_by_substring(['!!! kein Kursbetrieb !!!', 'XXX', '---'])
        self._drop_na()

        self.df.reset_index(drop=True, inplace=True)

        return self


if __name__ == '__main__':
    a = Path(r'.').glob(r'*.xlsx')
    assert len(a) == 1

    df = pd.read_excel(a[0], engine='openpyxl')
    b = Worker(df)
    df = b.df

    print(df.head())

    row_one = df.iloc[0]

