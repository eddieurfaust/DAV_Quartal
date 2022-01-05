import math
import tabula
import pandas as pd
from datetime import date, time, datetime

FILE = r'Kursplan Q1-2022.pdf'

# read PDF file
tables = tabula.read_pdf(FILE, pages="all")
tables2 = tables.copy()


def append_df(df, to_append) -> pd.DataFrame:
    # second frame: set header as new row at position 0
    df = df.copy()  # accumulate all data in this frame
    to_append = to_append.copy()
    df.loc[len(df)] = to_append.columns  # append header from new df (this is actually data)
    to_append.columns = df.columns  # update the header
    return pd.concat([df, to_append], ignore_index=True)


while len(tables) > 1:
    tables[0] = append_df(tables[0], tables[1])
    del tables[1]

assert len(tables) == 1
df = tables[0]
print(df.shape)


# HERE WE ARE DONE


# ALTERNATIVE
def concat_all(lst) -> pd.DataFrame:
    if len(lst) == 1:
        return lst[0]

    return concat_all([append_df(lst[0], lst[1]), *lst[2:]])
    # lst[0] = append_df(lst[0], lst[1])
    # del lst[1]


a = concat_all(tables2)

assert a.equals(df)


def update_date_str(inp: str) -> str:
    if type(inp) == str and inp.count('.') == 3:
        return '.'.join(inp.split('.')[:-1])

    return inp


# remove artifact when stitching the frames
a['Termin 1'] = a['Termin 1'].map(update_date_str)

# drop columns that are not used
a = a.drop(['Öffn.', 'Tag', 'Datum'], axis=1)

print(a.shape)

# drow rows
# a = a.dropna(subset=['Trainer 1'])
# a = a.dropna(subset=['Termin 1'])
a = a[a["Kurs"].str.contains("XXX|!!!|---") == False]


a = a.reset_index()

### hiermit kann man nun arbeiten
cur_row = a.loc[1]
# Trainer: Trainer 1
tr1 = cur_row['Trainer 1']
# Backup: Trainer 2
tr2 = cur_row['Trainer 2']
# Datum: Termin 1..4
def get_days(inp):
    lst = []
    for t in ['Termin 1', 'Termin 2', 'Termin 3', 'Termin 4']:
        if type(inp[t]) == str:
            a = datetime.strptime(inp[t], "%d.%m.%y").date()
            lst.append(a)

    return lst
ds = get_days(cur_row)
# Uhrzeit: Uhrzeit
def get_time(inp) -> time:
    begin, end = inp['Uhrzeit'].split('-')
    begin = datetime.strptime(begin.strip(), "%H:%M").time()
    end = datetime.strptime(end.strip(), "%H:%M").time()

    return begin, end
time_ = get_time(cur_row)
# Description: Kurs
desc = cur_row['Kurs'].replace(' - ', '-')

print(a.shape)
