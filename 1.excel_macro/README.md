
# 사용법
## Python 설치
- [설치링크](https://www.python.org/downloads/)
- 코드 테스트 : Python 3.10.4

## 가상환경 생성 및 패키지 설치
``` shell
# 경로 설정
cd {경로}

# 가상환경
python -m venv .venv
source .venv/bin/activate     # linux or mac
# .venv\Scripts\activate      # windows

# 패키지 설치
pip install -r requirements.txt
```
## 코드 실행
- weekly 작업 코드
    ```
    python excel_macro.py  -t weekly
    # --세부 경로 입력
    ```
- monthly 작업 코드
    ```
    python excel_macro.py  -t monthly
    # 세부 경로 입력
    ```

# 작업 구조
## 필요 패키지
- os : 디렉터리에 어떤 파일이 있는지 알아보고, 그 파일명을 가져옴
- re : 우편번호 가져오는 정규식에 필요
- warnings : 간단한 오류 무시
- pandas : 엑셀 파일을 읽고, 처리하고, 저장
- openpyxl : 엑셀 파일의 셀 서식을 가져옴
- datetime : 현재 날짜 가져오기

## 작업 Tree
```
.
├── Excel
│   ├── 결과
│   │   ├── 저널홍보제품목록_{yyyy-mm}_{yyyy-mm-dd}.xlsx
│   │   └── 저작권_{yyyy-mm}_{yyyy-mm-dd}.xlsx
│   └── 제품목록
│       ├── 2023-05
│       │   ├── file1.xlsx
│       │   ├── file2.xlsx
│       │   └── file2.xlsx
│       ├── 2023-06
│       │   ├── file1.xlsx
│       │   ├── file2.xlsx
│       │   └── file3.xlsx
│       └── ..
├── README.md
├── requirements.txt
└── excel_macro.py
```

## 코드 로직
### excel_macro.py [type : monthly]
```
input : 세부 파일 경로
output : 엑셀 파일 ['저널홍보제품목록_{year_month}_{today}.xlsx']
```
1. os로 해당 경로에 어떤 파일 목록이 있는지 가져온다.
2. 그 파일 목록을 읽는다.
3. 작업할 파일을 openpyxl로 읽는다.
 3-1. DataFrame으로 변경한다.
 3-2. 컬러를 파악할 파일을 workbook 형태로 넘겨준다
    -> 가져올 타겟 row 수를 체크한다.
4. 필요한 데이터를 가져온다.
 4-1. 필요한 columns, rows 만큼 가져온다.                      -> result_journal(저널복붙)
 4-2. 필요한 데이터를 변형해주고, 필요한 columns, rows 만큼 가져온다. -> result_distribution_list(배포처목록)
5. 파일을 각 시트에 저장한다.


### excel_macro_weekly.py
1. os로 해당 경로에 어떤 파일 목록이 있는지 가져온다.
2. 그 파일 목록을 읽는다.
3. 작업할 파일을 openpyxl로 읽는다.
 3-1. DataFrame으로 변경한다.
 3-2. 컬러를 파악할 파일을 workbook 형태로 넘겨준다
    -> 가져올 타겟 row 수를 체크한다.
4. 필요한 데이터를 가져온다.
5. 파일을 저장한다.
