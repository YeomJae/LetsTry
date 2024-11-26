import os
import pandas as pd
from tkinter import filedialog, messagebox, Tk
from openpyxl import load_workbook

# 시트 이름 정의
SHEET_NAMES = {
    "raw": "입력자료",
    "pivot_data": "피봇자료",
    "sort_data": "분류자료"
}

# 폴더 선택 GUI
root = Tk()
root.withdraw() # Tkinter 창 숨기기
input_folder_path = filedialog.askdirectory(parent=root, initialdir='D:/', title='폴더를 선택하세요')
if not input_folder_path:
    messagebox.showwarning("경고","폴더를 선택하지 않았습니다.")
    exit()

# 폴더 내 xls 파일 목록 확인
file_list = [f for f in os.listdir(input_folder_path) if f.endswith('.xls') or f.endswith('.xlsx')]

if not file_list:
    messagebox.showwarning("경고", "선택한 폴더에 .xls 또는 .xlsx 파일이 없습니다.")
    exit()

for file_name in file_list:
    try:
        # 파일 읽기
        file_path = os.path.join(input_folder_path, file_name)
        df_rawdata = pd.read_excel(file_path)

        # 피벗 테이블 생성
        df_piv = df_rawdata.pivot_table(index='품명', values='수량', aggfunc='sum', fill_value=0)\
        
        # 정렬된 데이터 생성
        df_sort = df_piv.sort_values(by='수량', ascending = False)

        # 결과 저장 경로 설정
        output_file_path = os.path.join(input_folder_path, f"processed_{file_name}")

        # 결과 저장
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            df_rawdata.to_excel(writer, sheet_name=SHEET_NAMES["raw"], index=False)
            df_piv.to_excel(writer, sheet_name=SHEET_NAMES["pivot_data"])
            df_sort.to_excel(writer, sheet_name=SHEET_NAMES["sort_data"])

        print(f"처리가 완료되었습니다: {output_file_path}")

    except Exception as e:
        print(f"파일 처리 중 오류 발생: {file_name} - {e}")