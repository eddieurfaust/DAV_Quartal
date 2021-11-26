import pandas as pd

from path import Path
from dataclasses import dataclass
from typing import Tuple, List
from datetime import datetime
from dateutil import parser
from tqdm import tqdm


@dataclass(frozen=True)
class Termin:
    subject: str
    start: datetime
    end: datetime
    description: str


class Worker:
    def __init__(self, df_: pd.DataFrame):
        self.df = df_
        self._clean_dataframe()

    def _drop(self, cols: list, rows: list):
        self.df.drop(cols, axis=1, inplace=True)
        self.df.drop(rows, axis=0, inplace=True)

        return self

    def _set_column_headers(self):
        self.df.columns = [e.strip() for e in self.df.iloc[0].to_list()]

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

    def _get_start_end_datetimes(self, idx: pd.Index):
        def get_start_end_datetime(date_: datetime.date, start_: datetime.time, end_: datetime.time) -> Tuple:
            ret_start = datetime.combine(date_.date(), parser.parse(start_).time())
            ret_end = datetime.combine(date_.date(), parser.parse(end_).time())

            return ret_start, ret_end

        list_of_classes = []

        for cur_termin_no in range(1, 5):
            cur_termin_str = f'Termin {cur_termin_no}'
            cur_date = self.df.loc[idx][cur_termin_str]

            if cur_date is None:
                # print(f'{cur_termin_str} -> None!')
                continue

            start, end = self.df.loc[idx]['Uhrzeit'].split(' - ')

            start_datetime, end_datetime = get_start_end_datetime(cur_date, start, end)

            list_of_classes.append((cur_termin_str, start_datetime, end_datetime))

        return list_of_classes

    def _get_subject(self, idx: pd.Index) -> str:
        return self.df.loc[idx]["Kurs"].replace(" -", "-").replace("- ", "-")

    def _get_description(self, idx: pd.Index) -> str:
        names_only = self.df.loc[idx].drop(
            ['Öffn.', 'Tag', 'Datum', 'Kurs', 'Uhrzeit', 'Termin 1', 'Termin 2', 'Termin 3', 'Termin 4'])
        list_of_trainers = names_only.dropna().to_list()
        list_of_trainers = sorted([n.capitalize() for n in list_of_trainers])

        return f'[Potentially assigned trainers] {" | ".join(list_of_trainers)}'

    def get_list_of_class_tuples(self):
        all_classes: [List[Termin]] = []
        for idx, row in tqdm(self.df.iterrows(), total=len(self.df)):
            cur_subject = self._get_subject(idx)
            cur_list_of_subclasses = self._get_start_end_datetimes(idx)
            cur_description = self._get_description(idx)

            for sub_class_str, start_datetime, end_datetime in cur_list_of_subclasses:
                all_classes.append(Termin(
                    subject=f'{cur_subject} - {sub_class_str}',
                    start=start_datetime,
                    end=end_datetime,
                    description=cur_description
                ))

        return all_classes


def to_gcal_csv(inp: List[Termin]) -> str:
    # strftime: https://www.programiz.com/python-programming/datetime/strftime
    out = 'Subject,Start Date,Start Time,End Time,Description,All Day Event'
    out += '\n'

    cur_row = inp[0]

    for cur_row in inp:
        row_elems = [
            cur_row.subject,  # Subject
            cur_row.start.strftime('%m/%d/%Y'),  # Start Date
            cur_row.start.strftime('%I:%M %p'),  # Start Time
            cur_row.end.strftime('%I:%M %p'),  # End Time
            cur_row.description,  # Description
            'False'  # All Day Event
        ]

        out += ','.join(row_elems)
        out += '\n'

    return out


def create_google_cal_file(inp: Path, out: Path = 'out.csv') -> None:
    df = pd.read_excel(str(inp), engine='openpyxl')
    worker = Worker(df)

    all_classes = worker.get_list_of_class_tuples()
    all_classes = [c for c in all_classes if 'Frank' in c.description]

    csv = to_gcal_csv(all_classes)

    with open(out, 'wt') as f:
        f.write(csv)

    print(f'wrote {len(all_classes)} classes')


if __name__ == '__main__':
    a = Path(r'.').glob(r'*.xlsx')
    assert len(a) == 1

    create_google_cal_file(a[0], Path('out.csv'))
