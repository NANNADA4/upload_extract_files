"""
폴더를 순회하면서 파일을 찾고 파일명을 변경하고 폴더를 생성합니다.
"""


import os
import shutil
import pandas as pd
from natsort import natsorted

from module.process_log import create_log
from module.tmp.tmp import tmp


def process_folder(input_num) -> bool:
    """전달받은 input_num을 토대로 알맞은 함수로 반환합니다."""
    input_path = input("입력 폴더의 경로를 입력하세요 : ")
    input_path = os.path.join('\\\\?\\', input_path)

    if not os.path.isdir(input_path):
        print("\n!!!!!입력 폴더의 경로를 다시 한 번 확인하세요!!!!!\n")
        return False

    match input_num:
        case '1':
            create_excel_path = input("엑셀 파일을 저장할 경로를 입력하세요 : ")
            process_create_excel(input_path, create_excel_path)
        case '2':
            base_excel_path = input("1번에서 실행한 결과의 엑셀파일 경로를 입력하세요 : ")
            attach_excel_path = input("별도제출자료를 정리한 엑셀파일 경로를 입력하세요 : ")
            process_merge_attach(base_excel_path, attach_excel_path)
        case '3':
            rename_output_path = input("경로 및 이름을 변경한 파일을 복사할 폴더 경로를 입력하세요 : ")
            rename_excel_path = input("엑셀 파일의 경로를 입력하세요 (엑셀이 종료되었는지 확인하세요) : ")
            process_rename(input_path, rename_output_path, rename_excel_path)

    return True


def process_merge_attach(base_excel_path, attach_excel_path):
    """1번에서 제작한 엑셀파일과 별도제출자료를 정리한 엑셀파일을 병합합니다"""
    #! TODO => Remove tmp()
    tmp(base_excel_path, attach_excel_path)


def process_create_excel(input_path, excel_path):
    """BOOK_ID, SEQNO가 매겨진 엑셀파일을 생성합니다"""
    #! TODO => Remove tmp()
    tmp(input_path, excel_path)


def process_rename(input_path, output_path, excel_path):
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
                create_log(os.path.join(root, file))
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
