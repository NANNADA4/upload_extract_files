"""BOOK_ID, SEQNO가 매겨진 엑셀파일을 생성합니다"""


import os


from module.create_excel.create_excel import create_excel


def process_create(input_path, excel_path):
    """BOOK_ID, SEQNO가 매겨진 엑셀파일을 생성합니다"""
    try:
        create_excel(input_path, excel_path)
        print("\n>>> 엑셀 파일이 정상적으로 생성되었습니다 <<<\n")
    except Exception as e:  # pylint: disable=W0718
        print(f"\n!!!!!!!{e}!!!!!!!\n")
        os.remove(excel_path)
