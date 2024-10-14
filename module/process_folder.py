"""
폴더를 순회하면서 파일을 찾고 파일명을 변경하고 폴더를 생성합니다.
"""


import os
import shutil
import pandas as pd
from natsort import natsorted

from module.write_log import write_log


def process_folder(excel_path, input_path, output_path):
    """폴더를 순회하면서 엑셀 파일 내부 파일명을 변경하고 새로운 폴더로 이동합니다."""
    df = pd.read_excel(excel_path, engine='openpyxl')

    if not all(col in df.columns for col in ['실제 파일명', 'FILE_NAME', 'FILE_PATH']):
        print("엑셀 읽기 오류! : 엑셀 파일에 필요한 열이 없습니다.")
        return

    df['FILE_PATH'] = df['FILE_PATH'].astype(str)

    for root, _, files in os.walk(input_path):
        for file in natsorted(files):
            matching_row = df[df['실제 파일명'] == file]

            if matching_row.empty:
                write_log(os.path.join(output_path, 'log.txt'),
                          os.path.join(root, file))
                continue

            for index, row in matching_row.iterrows():
                actual_file_path = os.path.join(root, file)

                file_path_row = f"/inspection/reqdoc/2023/{row['BOOK_ID']}/" + \
                    f"{row['FILE_NAME']}"
                real_file_path = os.path.join(
                    "inspection", "reqdoc", "2023", str(row['BOOK_ID']), str(row['FILE_NAME']))
                target_file_path = os.path.join(output_path, real_file_path)

                target_folder = os.path.dirname(target_file_path)
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)

                if os.path.exists(actual_file_path):
                    shutil.copy(actual_file_path, target_file_path)

                df.at[index, 'FILE_PATH'] = file_path_row
                df.at[index, 'FILE_NAME'] = row['FILE_NAME']

    print("\n~~~엑셀 파일 수정중입니다~~~")

    try:
        df.to_excel(excel_path, index=False, engine='openpyxl')
    except PermissionError:
        print("엑셀 수정 실패! : 엑셀 파일이 열려있는 경우, 닫고 다시 실행하세요")
        return
