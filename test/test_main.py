from pathlib import Path

import pytest

from main import create_google_cal_file
import datetime


# setup test data
_14 = {'Uhrzeit': '10:00 - 22:00', 'Tag': datetime.datetime(2021, 10, 5, 0, 0), 'Datum': datetime.datetime(2022, 7, 12, 0, 0), 'Kurs': 'Offener Klettertreff', 'Zeit': '19:00 - 21:00', 'Termin 1': datetime.datetime(2022, 7, 12, 0, 0), 'Termin 2': None, 'Termin 3': None, 'Termin 4': None, 'Leonard': None, 'Marco': None, 'Manuela': None, 'Sarah': 'Sarah', 'Micha': None, 'Cornelia': None, 'Andre': None, 'Kerstin': None, 'Stefan G.': None, 'Heidi': None, 'Frank': None}
_26 = {'Uhrzeit': '10:00 - 20:00', 'Tag': datetime.datetime(2021, 10, 2, 0, 0), 'Datum': datetime.datetime(2022, 7, 23, 0, 0), 'Kurs': 'Vorstiegs - Kurs', 'Zeit': '10:30 - 14:30', 'Termin 1': datetime.datetime(2022, 7, 23, 0, 0), 'Termin 2': datetime.datetime(2022, 7, 24, 0, 0), 'Termin 3': None, 'Termin 4': None, 'Leonard': None, 'Marco': 'Marco', 'Manuela': 'Manuela', 'Sarah': None, 'Micha': 'Micha', 'Cornelia': None, 'Andre': None, 'Kerstin': None, 'Stefan G.': None, 'Heidi': '? Hospitation Heidi', 'Frank': None}
_91 = {'Uhrzeit': '14:00 - 22:00', 'Tag': datetime.datetime(2021, 10, 4, 0, 0), 'Datum': datetime.datetime(2022, 9, 12, 0, 0), 'Kurs': 'Aufbau - Kurs für Einsteiger', 'Zeit': '18:30 - 20:30', 'Termin 1': datetime.datetime(2022, 9, 12, 0, 0), 'Termin 2': datetime.datetime(2022, 9, 19, 0, 0), 'Termin 3': datetime.datetime(2022, 9, 26, 0, 0), 'Termin 4': datetime.datetime(2022, 10, 3, 0, 0), 'Leonard': None, 'Marco': None, 'Manuela': None, 'Sarah': None, 'Micha': None, 'Cornelia': 'Cornelia', 'Andre': None, 'Kerstin': None, 'Stefan G.': None, 'Heidi': None, 'Frank': 'Frank (Hospi)'}


# overarching test of the whole process input => output
@pytest.mark.skip(reason='temporarily disable to check coverage')
def test_create_google_cal_file():
    out_path = Path(r'out.ics')
    expected_path = Path('expected_data/out.ics')
    create_google_cal_file(Path(r'test_data/Vorplan Plan Q3-22 v4 - Für die Trainer.xlsx'), out_path)

    try:
        with open(expected_path, 'rt', encoding='UTF-8') as expected:
            with open(out_path, 'rt', encoding='UTF-8') as test:
                for exp, tst in zip(expected.readlines(), test.readlines()):
                    assert exp == tst

    finally:
        out_path.unlink(missing_ok=True)
