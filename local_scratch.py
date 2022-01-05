import tabula
import pandas as pd

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

# second frame: set same column names (necessary?)

# concat 2 frames
# a = pd.concat([tables[0], tmp])


print(tables)

# Idee: nur mit puren Funktionen arbeiten
#  - das ist hier zwar nicht sehr sinnig, da die Dataframes immer wieder kopiert werden

# tables zu einem sauberen großen Frame zusammen ziehen
#  - f([table, table]) -> table ()
#    - soll 2 tables zusammen führen und dabei den Header der ersten übernehmen
#    - Annahme: erster table hat den header

# filtern/aussortieren: Spalten Öffn., Tag, Datun können komplett raus, Zeilen mit XXX oder !!! im Kurs können auch raus

# Trainer: Trainer 1
# Backup: Trainer 2
# Datum: Termin 1..4
# Uhrzeit: Uhrzeit
# Description: Kurs
