from pathlib import Path

from main import create_google_cal_file


def test_create_google_cal_file():
    out_path = Path('test_out.csv')
    expected_path = Path('expected_out.csv')
    create_google_cal_file(Path(r'D:\MMS\Python\DAV_Quartal\Vorplan Plan Q1-22 v4 - Für die Trainer.xlsx'), out_path)

    try:
        with open(expected_path, 'rt') as expected:
            with open(out_path, 'rt') as test:
                for exp, tst in zip(expected.readlines(), test.readlines()):
                    assert exp == tst

    finally:
        out_path.unlink(missing_ok=True)
