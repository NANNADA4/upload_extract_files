"""1번에서 제작한 엑셀파일과 별도제출자료를 정리한 엑셀파일을 병합합니다"""


from module.excel.compare_excel import compare_excel
from module.excel.load_excel import load_excel


def process_merge(base_excel_path, attach_excel_path):
    """1번에서 제작한 엑셀파일과 별도제출자료를 정리한 엑셀파일을 병합합니다"""
    try:
        compare_excel(load_excel(base_excel_path), load_excel(
            attach_excel_path)).save(base_excel_path)
        print("\n=> 엑셀 파일이 정상적으로 병합되었습니다\n")
    except Exception as e:  # pylint: disable=W0718
        print(e)
