"""BOOK_ID, SEQNO가 매겨진 엑셀파일을 생성합니다"""


from module.create_excel.create_excel import create_excel


def process_create(input_path, excel_path):
    """BOOK_ID, SEQNO가 매겨진 엑셀파일을 생성합니다"""
    create_excel(input_path, excel_path)
