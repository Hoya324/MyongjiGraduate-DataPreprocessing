import pandas as pd
import warnings
from glob import glob
import os

warnings.filterwarnings(action='ignore')

# 결과를 저장할 데이터프레임 생성
total_Y_major = pd.DataFrame(columns=['code', 'name', 'credit', 'isRevoked', 'duplicatedCode'])

# 엑셀 파일이 있는 디렉토리 경로
excel_files_directory = 'data/자연캠퍼스/*/'

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
    total_Y_major = pd.concat([total_Y_major, data], ignore_index=True)

# 결과 데이터프레임을 새로운 엑셀 파일로 저장
total_Y_major.to_excel('data/자연캠퍼스_통합교과코드.xlsx', index=False)
