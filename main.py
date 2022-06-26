import pandas as pd

from path import Path
import toml
from dataclasses import dataclass
import re
from typing import Tuple, List, Optional, Union
from datetime import datetime
from dateutil import parser
from icalendar import Calendar, Event
from tqdm import tqdm

from Logger import logger as log


LOCATION = 'GriffReich - Kletterzentrum des DAV Hannover, Peiner Str. 28, 30519 Hannover, Deutschland'


@dataclass(frozen=True)
class Termin:
    subject: str
    start: datetime
    end: datetime
    description: str


class Defaults:
    def __init__(self, toml_path: Path = Path(r'defaults.toml')):
        self.defaults = {
            'skip_lines': 1,
            'delete_courses_named': ['!!! kein Kursbetrieb !!!', 'XXX', '---'],
            'bullet_point': ' • ',
            'highlight_chars': '<-',
        }

        self.toml_dict = self._load_defaults_toml(toml_path)

    @staticmethod
    def _load_defaults_toml(toml_path: Path) -> Optional[dict]:
        try:
            input_file_name = str(toml_path)
            with open(input_file_name, encoding='UTF-8') as toml_file:
                toml_dict = toml.load(toml_file)
            log.info(f'successfully read defaults file: {toml_path}')
            log.debug(f'defaults: {toml_dict}')

            return toml_dict
        except:
            log.info(f'could not read defaults from {toml_path}')

        return None

    def get_parameter(self, inp: str):
        defaults = {
            'skip_lines': 1,
            'delete_courses_named': ['!!! kein Kursbetrieb !!!', 'XXX', '---'],
            'bullet_point': ' • ',
            'highlight_chars': '<-',
        }

        if self.toml_dict is not None and self.toml_dict.get(inp) is not None:
            # log.debug(f'getting from config file: [{inp}] : [{self.toml_dict.get(inp)}]')
            return self.toml_dict.get(inp)

        # log.debug(f'getting from build-in defaults: [{inp}] : [{defaults.get(inp)}]')
        return defaults.get(inp)


class Worker:
    def __init__(self, df_: Union[pd.DataFrame, None] = None):
        self.defaults = Defaults()

        self.df = df_
        if self.df is not None:
            self._clean_dataframe()

    def _drop(self, cols: list, rows: list):
        self.df.drop(cols, axis=1, inplace=True, errors='ignore')
        self.df.drop(rows, axis=0, inplace=True, errors='ignore')

        return self

    def _set_column_headers(self):
        self.df.columns = [e.strip() for e in self.df.iloc[0].to_list()]

        self._drop(cols=[], rows=[self.df.index[0]])

        return self

    def _replace_nan_with_none(self):
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
        # drop (empty) dummy cols Name 01, Name 02, ..., Name 20
        # drop first x rows
        self._drop(cols=[f'Name {x:02}' for x in range(21)], rows=list(range(1 + self.defaults.get_parameter('skip_lines'))))
        self._drop_na()
        self._set_column_headers()
        self._replace_nan_with_none()
        self._remove_rows_by_substring(self.defaults.get_parameter('delete_courses_named'))
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

        ret = ''
        sep = self.defaults.get_parameter('bullet_point')
        break_sep = f'\n{sep}'
        if list_of_trainers:
            ret += f'[Registered potential instructors:]\n{sep}{break_sep.join(list_of_trainers)}'
        else:
            ret += f'So far no one wants to teach this course.'

        termin_string = '\n\n[Termine:]'
        termine = self.df.iloc[idx][[x for x in self.df.columns if 'Termin' in x]]
        for name, dt in zip(termine.index.to_list(), termine.to_list()):
            if dt is None:
                continue

            termin_string += f'\n{sep}{name} - {dt.date()}'

        ret += termin_string

        return ret

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

    def highlight_termin(row) -> str:
        found = re.search('Termin \d', row.subject)
        if not found:
            return row.description

        out = ''
        lines = row.description.split('\n')

        highlight_chars = Defaults().get_parameter('highlight_chars')

        for line in lines:
            if found.group(0) in line:
                line += ' ' + highlight_chars
            out += line + '\n'

        return out.rstrip('\n')

    for row in inp:
        event = Event()
        event.add('summary', f'🧗 {row.subject}')
        event.add('dtstart', row.start)
        event.add('dtend', row.end)
        event.add('location', LOCATION)
        event.add('description', highlight_termin(row))  # returns potentially altered `row.description`

        calendar_events.append(event)

    cal = Calendar()
    [cal.add_component(event) for event in calendar_events]

    with open(str(out_filename), 'wb') as f:
        f.write(cal.to_ical())

    print(f'wrote {len(inp)} classes')


def create_google_cal_file(inp: Path, out: Path = 'automatic_out.ics') -> None:
    df = pd.read_excel(str(inp), engine='openpyxl')
    worker = Worker(df)

    all_classes = worker.get_list_of_class_tuples()

    # MAYBE FILTER HERE
    # all_classes = [c for c in all_classes if 'Frank' in c.description]

    to_gcal_ical(all_classes, out)


def vorplan_worker(path: Path):
    files = path.glob(r'*.xlsx')
    files = [x for x in files if not str(x.name).startswith('~')]

    assert len(files) == 1, f'files found: {files}'

    create_google_cal_file(files[0], Path('out.ics'))


if __name__ == '__main__':
    vorplan_worker(Path(r'./Q3/Vorplan'))
