import os
import pandas as pd
import util as ut

current_dir = ut.exedir('py')

def xlsxProcess(input_file_name, output_file_name):
    """
    엑셀 파일을 읽고, 지정된 컬럼을 제거한 후 진행상태별로 데이터를 분류하고 요청일자로 정렬하여
    새로운 엑셀 파일로 저장하는 함수입니다.
    """
    # 현재 작업 디렉토리의 'download' 폴더에서 파일 경로 설정
    input_file_path = os.path.join(current_dir, 'download', input_file_name)
    output_file_path = os.path.join(current_dir, 'result', output_file_name)

    # 엑셀 파일 로드
    excel_data = pd.ExcelFile(input_file_path)
    df = excel_data.parse('Sheet')  # 'Sheet'는 시트 이름입니다.

    # 컬럼 제거
    columns_to_remove = ['선택', 'No', '세부구분', '구매진행상세', '접수자']
    df_cleaned = df.drop(columns=columns_to_remove)

    # 진행상태별로 데이터 분류 및 요청일자 기준 정렬
    grouped_data = {status: group.sort_values(by='요청일자') for status, group in df_cleaned.groupby('진행상태')}

    # 새로운 엑셀 파일 생성
    with pd.ExcelWriter(output_file_path) as writer:
        for status, group in grouped_data.items():
            group.to_excel(writer, sheet_name=str(status), index=False)

    print(f"작업 완료. 파일이 '{output_file_path}'에 저장되었습니다.")