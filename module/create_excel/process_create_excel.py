"""BOOK_ID, SEQNO가 매겨진 엑셀파일을 생성합니다"""


from module.tmp.tmp import tmp


def process_create(input_path, excel_path):
    """BOOK_ID, SEQNO가 매겨진 엑셀파일을 생성합니다"""
    #! TODO => Remove tmp()
    tmp(input_path, excel_path)
