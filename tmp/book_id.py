"""
BOOK_ID와 FILE_NAME열을 읽어 폴더를 만들어 복사하는 스크립트입니다.
"""


import os
import shutil
import pandas as pd


def create_folders_and_copy_files(src_folder, excel_file, dest_folder):
    """폴더 순회를 하며 FILE_NAME과 동일한 파일이 발견될 경우, 생성된 BOOK_ID폴더로 복사합니다"""
    df = pd.read_excel(excel_file)

    for root, _, files in os.walk(src_folder):
        for file in files:
            matched_row = df[df['FILE_NAME'] == file]
            if matched_row.empty:
                print(f"{file}을 찾을 수 없습니다")
                continue

            book_id = matched_row.iloc[0]['BOOK_ID']
            new_folder_path = os.path.join(dest_folder, str(book_id))

            os.makedirs(new_folder_path, exist_ok=True)

            src_file_path = os.path.join(root, file)
            dest_file_path = os.path.join(new_folder_path, file)
            shutil.copy(src_file_path, dest_file_path)
            print(f"{file} 복사 성공")


def main():
    """메인함수"""
    print("-"*24)
    print("\n>>>>>>파일 복사<<<<<<\n")
    print("***'FILE_NAME', 'BOOK_ID' 열이 존재하는지 확인하세요***\n")
    print("-"*24)
    src_folder = input("원본 폴더 경로를 입력하세요 : ")
    excel_file = input("엑셀 파일 경로를 입력하세요(확장자 포함) : ")
    dest_folder = input("목적지 폴더 경로를 입력하세요 : ")
    create_folders_and_copy_files(src_folder, excel_file, dest_folder)

    if __name__ == "__main__":
        main()
