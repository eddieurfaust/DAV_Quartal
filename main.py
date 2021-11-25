from path import Path
import pandas as pd


if __name__ == '__main__':
    a = Path(r'.').glob(r'*.xlsx')
    assert len(a) == 1

    xlsx = pd.read_excel(a[0])

    print(xlsx)

