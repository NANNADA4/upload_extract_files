"""1번에서 제작한 엑셀파일과 별도제출자료를 정리한 엑셀파일을 병합합니다"""


from module.tmp.tmp import tmp


def process_merge(base_excel_path, attach_excel_path):
    """1번에서 제작한 엑셀파일과 별도제출자료를 정리한 엑셀파일을 병합합니다"""
    #! TODO => Remove tmp()
    tmp(base_excel_path, attach_excel_path)
