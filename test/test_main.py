from pathlib import Path

from main import create_google_cal_file


# overarching test of the whole process input => output
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
