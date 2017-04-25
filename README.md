# excel-compare
The excel compare package for girlfriend (kkk)

## How to use

python 3.5.3 이 설치되어 있어야 합니다.

`run.py` 를 아래와 같이 작성합니다

```py
from excel_compare import ExcelMeta, ExcelCompare
excel_meta_a=ExcelMeta(os.path.join('src', 'a.xlsx'), 'sheet_a', [])
excel_meta_b=ExcelMeta(os.path.join('src', 'b.xlsx'), 'sheet_b', [])


excel_compare = ExcelCompare(excel_meta_a, excel_meta_b, os.path.join('out', '중복.xlsx'))
excel_compare.analyze()
```

이후에 아래와 같이 의존성 라이브러리를 설치하고 실행합니다.

```sh
pip install -r requirements.txt
python3 run.py
```
