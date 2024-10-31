"""
폴더를 순회하면서 파일을 찾고 파일명을 변경하고 폴더를 생성합니다.
"""


import os


from module.create_excel.process_create_excel import process_create
from module.merge_attach.process_merge_attach import process_merge
from module.rename_files.process_rename_files import process_rename


def process_folder(input_num) -> bool:
    """전달받은 input_num을 토대로 알맞은 함수로 반환합니다."""
    if input_num in ['1', '3']:
        input_path = input("입력 폴더의 경로를 입력하세요 : ")
        input_path = os.path.join('\\\\?\\', input_path)

        if not os.path.isdir(input_path):
            print("\n====입력 폴더의 경로를 다시 한 번 확인하세요====\n")
            return False

    match input_num:
        case '1':
            while True:
                create_excel_path = input("엑셀 파일을 저장할 경로를 입력하세요 : ")
                if not create_excel_path.lower().endswith('.xlsx'):
                    print("====엑셀 파일 확장자를 입력했는지 확인해주세요====")
                process_create(input_path, create_excel_path)
                break
        case '2':
            while True:
                base_excel_path = input("1번에서 실행한 결과의 엑셀파일 경로를 입력하세요 : ")
                attach_excel_path = input("별도제출자료를 정리한 엑셀파일 경로를 입력하세요 : ")
                if os.path.exists(base_excel_path) and os.path.exists(attach_excel_path):
                    process_merge(base_excel_path, attach_excel_path)
                    break
                print("\n=> 엑셀 파일 경로를 한번 더 확인해주세요")
        case '3':
            rename_output_path = input("경로 및 이름을 변경한 파일을 복사할 폴더 경로를 입력하세요 : ")
            rename_excel_path = input("엑셀 파일의 경로를 입력하세요 (엑셀이 종료되었는지 확인하세요) : ")
            process_rename(input_path, rename_output_path, rename_excel_path)

    return True
