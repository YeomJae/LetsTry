import os
import pandas as pd
from tkinter import filedialog, messagebox, Tk
from openpyxl import load_workbook
import util as ut

# 시트 이름 정의
SHEET_NAMES = {
    "raw": "구매진행관리",
    "pivot_data": "피봇자료",
    "sort_data": "분류자료"
}

def xlsxprocess(file_name)
    current_dir = ut.exedir('py')
    file_path = os.path.join(current_dir, "download")
    
    df_rawdata = pd.read_excel(file_path)
    df_sort = df_rawdata(by='진행현황', ascending = False)

    # 결과 저장 경로 설정
    output_file_path = os.path.join(input_folder_path, f"processed_{file_name}")

    # 결과 저장
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        df_rawdata.to_excel(writer, sheet_name=SHEET_NAMES["raw"], index=False)
        df_piv.to_excel(writer, sheet_name=SHEET_NAMES["pivot_data"])
        df_sort.to_excel(writer, sheet_name=SHEET_NAMES["sort_data"])