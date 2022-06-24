import pandas as pd

from path import Path
from dataclasses import dataclass
from typing import Tuple, List
from datetime import datetime
from dateutil import parser
from icalendar import Calendar, Event
from tqdm import tqdm


@dataclass(frozen=True)
class Termin:
    subject: str
    start: datetime
    end: datetime
    description: str


def get_parameter(inp: str):
    if inp == 'skip_lines':
        return 1
    if inp == 'delete_courses_named':
        return ['!!! kein Kursbetrieb !!!', 'XXX', '---']
    return None


class Worker:
    def __init__(self, df_: pd.DataFrame):
        self.df = df_
        self._clean_dataframe()

    def _drop(self, cols: list, rows: list):
        self.df.drop(cols, axis=1, inplace=True, errors='ignore')
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
        # self._drop(cols=['Unnamed: 0', 'None.7'], rows=[0, 1])  # cols würden auch durch self.drop_na() weg fallen?
        self._drop(cols=[f'Name {x:02}' for x in range(21)], rows=range(1 + get_parameter('skip_lines')))
        self._drop_na()
        self._set_column_headers()
        self._replace_NaN_with_None()
        self._remove_rows_by_substring(get_parameter('delete_courses_named'))
        self._drop_na()

        self.df.reset_index(drop=True, inplace=True)

        return self

    def _get_start_end_datetimes(self, idx: pd.Index):
        """
        Returns the name, start and end time for all courses defined in the row at the given index.
        :param idx: index of the dataframe managed by the class
        :return: list containing zero to four courses which are extracted from one row
        """
        def get_start_end_datetime(date_: datetime.date, start_: datetime.time, end_: datetime.time) -> Tuple:
            ret_start = datetime.combine(date_.date(), parser.parse(start_).time())
            ret_end = datetime.combine(date_.date(), parser.parse(end_).time())

            return ret_start, ret_end

        list_of_classes = []

        for cur_termin_no in range(1, 5):
            cur_termin_str = f'Termin {cur_termin_no}'
            cur_date = self.df.loc[idx][cur_termin_str]  # this is a datetime object, we only need .date()

            if cur_date is None:
                continue

            start, end = self.df.loc[idx]['Zeit'].split(' - ')  # again, datetime objects, now we only need .time()

            start_datetime, end_datetime = get_start_end_datetime(cur_date, start, end)

            list_of_classes.append((cur_termin_str, start_datetime, end_datetime))

        return list_of_classes

    def _get_subject(self, idx: pd.Index) -> str:
        return self.df.loc[idx]["Kurs"].replace(" -", "-").replace("- ", "-")

    def _get_description(self, idx: pd.Index) -> str:
        names_only = self.df.loc[idx].drop(
            ['Öffn.', 'Tag', 'Datum', 'Kurs', 'Uhrzeit', 'Zeit', 'Termin 1', 'Termin 2', 'Termin 3', 'Termin 4'],
            errors='ignore')
        list_of_trainers = names_only.dropna().to_list()
        list_of_trainers = sorted([n.capitalize() for n in list_of_trainers])

        if list_of_trainers:
            return f'[Registered potential instructors:] {" | ".join(list_of_trainers)}'
        else:
            return f'So far no one wants to teach this course.'

    def get_list_of_class_tuples(self):
        all_classes: [List[Termin]] = []
        for idx, row in tqdm(self.df.iterrows(), total=len(self.df)):
            cur_subject = self._get_subject(idx)
            cur_list_of_subclasses = self._get_start_end_datetimes(idx)
            cur_description = self._get_description(idx)

            for sub_class_str, start_datetime, end_datetime in cur_list_of_subclasses:
                subject = f'{cur_subject}'
                if len(cur_list_of_subclasses) > 1:
                    subject += f' - {sub_class_str}'
                all_classes.append(Termin(
                    subject=subject,
                    start=start_datetime,
                    end=end_datetime,
                    description=cur_description
                ))

        return all_classes


def to_gcal_ical(inp: List[Termin], out_filename: str) -> str:
    calendar_events: List = []
    for row in inp:
        event = Event()
        event.add('summary', f'🧗 {row.subject}')
        event.add('dtstart', row.start)
        event.add('dtend', row.end)
        event.add('description', row.description)

        calendar_events.append(event)

    cal = Calendar()
    [cal.add_component(event) for event in calendar_events]

    with open(str(f'{out_filename}.ics'), 'wb') as f:
        f.write(cal.to_ical())

    print(f'wrote {len(inp)} classes')


def create_google_cal_file(inp: Path, out: Path = 'expected_out.csv') -> None:
    df = pd.read_excel(str(inp), engine='openpyxl')
    worker = Worker(df)

    all_classes = worker.get_list_of_class_tuples()

    # MAYBE FILTER HERE
    # all_classes = [c for c in all_classes if 'Frank' in c.description]

    to_gcal_ical(all_classes, out)


if __name__ == '__main__':
    a = Path(r'./Q3/Vorplan').glob(r'*.xlsx')
    a = [x for x in a if not str(x.name).startswith('~')]
    assert len(a) == 1, f'files found: {a}'

    create_google_cal_file(a[0], Path('out'))
