import pandas as pd
import warnings
from glob import glob
import os

warnings.filterwarnings(action='ignore')

# 결과를 저장할 데이터프레임 생성
total_S = pd.DataFrame(columns=['code', 'name', 'credit', 'isRevoked', 'duplicatedCode'])

# 엑셀 파일이 있는 디렉토리 경로
excel_files_directory = 'data/인문캠퍼스_및_교양/*/'

# 디렉토리 내의 모든 엑셀 파일에 대해 반복
for file_path in glob(os.path.join(excel_files_directory, '*.xlsx')):

    print(file_path)
    # 엑셀 파일에서 필요한 컬럼만 읽어오기
    data = pd.read_excel(file_path, usecols=['교과목명(국문)', '학점수', '교과코드', '중복코드', '폐지일자'])

    # 컬럼명 변경
    data.columns = ['code', 'name', 'credit', 'isRevoked', 'duplicatedCode']

    # NaN 값이 있는 경우 0으로, 없는 경우 1로 대체
    data['isRevoked'] = (~data['isRevoked'].isna()).astype(int)

    # 데이터를 결과 데이터프레임에 추가
    total_S = pd.concat([total_S, data], ignore_index=True)

# 결과 데이터프레임을 새로운 엑셀 파일로 저장
total_S.to_excel('data/인문캠퍼스_및_교양_통합교과코드.xlsx', index=False)

previous_excel = pd.read_excel('data/기존통합교과코드.xlsx')

# 'code'를 문자열로 변환 후 기준으로 오름차순 정렬
previous_excel['code'] = previous_excel['code'].astype(str)
previous_excel = previous_excel.sort_values(by='code')

new_excel = pd.read_excel('data/인문캠퍼스_및_교양_통합교과코드.xlsx')

# 'code'를 문자열로 변환 후 기준으로 오름차순 정렬
new_excel['code'] = new_excel['code'].astype(str)
new_excel = new_excel.sort_values(by='code')

# 정렬 후 이전과 변경된 값 비교
new_excel = new_excel.reindex(columns=previous_excel.columns, index=previous_excel.index)
is_changed = (previous_excel != new_excel).any(axis=1)

# isNew, isDelete 컬럼 추가하고 변경된 값에 'O' 추가
previous_excel['isNew'] = ''
previous_excel['isDelete'] = ''

# 새로운 데이터에 대해 isNew 컬럼에 'O' 추가
new_rows = previous_excel[~previous_excel['code'].isin(new_excel['code'])].index
previous_excel.loc[new_rows, 'isNew'] = 'O'

# 삭제된 데이터에 대해 isDelete 컬럼에 'O' 추가
deleted_rows = new_excel[~new_excel['code'].isin(previous_excel['code'])].index
previous_excel.loc[deleted_rows, 'isDelete'] = 'O'

# 새로운 데이터를 previous_excel에 추가
previous_excel = pd.concat([previous_excel, new_excel[~new_excel['code'].isin(previous_excel['code'])]])

previous_excel.to_excel('data/기존통합교과코드_수정.xlsx', index=False)
