import pandas as pd
import warnings
from glob import glob
import os

warnings.filterwarnings(action='ignore')

# 결과를 저장할 데이터프레임 생성
total_major = pd.DataFrame(columns=['name', 'credit', 'code', 'duplicatedCode'])

# 엑셀 파일이 있는 디렉토리 경로
excel_files_directory = 'data/전공/*/'

# 디렉토리 내의 모든 엑셀 파일에 대해 반복
for file_path in glob(os.path.join(excel_files_directory, '*.xlsx')):

    print(file_path)
    # 엑셀 파일에서 필요한 컬럼만 읽어오기
    data = pd.read_excel(file_path, usecols=['교과목명(국문)', '학점수', '교과코드', '중복코드'])

    # 컬럼명 변경
    data.columns = ['name', 'credit', 'code', 'duplicatedCode']
    data['id'] = ''
    # abolition 컬럼 추가
    data['abolition'] = ''

    # 데이터를 결과 데이터프레임에 추가
    total_major = pd.concat([total_major, data], ignore_index=True)

# 결과 데이터프레임을 새로운 엑셀 파일로 저장
total_major.to_excel('data/새로운통합교과코드.xlsx', index=False)

previous_excel = pd.read_excel('data/기존통합교과코드.xlsx')

new_excel = pd.read_excel('data/새로운통합교과코드.xlsx')

previous_excel = previous_excel.reindex(columns=new_excel.columns, index=new_excel.index)

# 변경된 값을 previous_excel에 반영하고 isChange 컬럼 추가
previous_excel['isChange'] = ''
for column in ['code']:
    changed_rows = new_excel[previous_excel[column] != new_excel[column]].index
    previous_excel.loc[changed_rows, column] = new_excel.loc[changed_rows, column]
    previous_excel.loc[changed_rows, 'isChange'] = 'O'


previous_excel.to_excel('data/기존통합교과코드_수정.xlsx', index=False)

#
# TODO: 아직 기존의 데이터와 새로운 데이터의 수정값이 적용되지 않음
# 1. 교양과목도 새로운통합교과코드 파일에 추가
# 2. 수정된 데이터의 행을 정렬?이든 올바른 값을 비교해서 isChange에 확인할 수 있도록
#