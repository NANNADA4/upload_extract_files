"""
두 엑셀 파일을 비교합니다.
"""

import os
from openpyxl import Workbook, load_workbook


def add_pdf_answer(excel_1: Workbook, excel_2: Workbook) -> Workbook:
    """두 엑셀 파일을 비교하여 값이 같을 경우 PDF상 답변을 추가합니다"""
    ws1 = excel_1.active  # 1번 스크립트에서 생성한 엑셀파일
    ws2 = excel_2.worksheets[1]
    excel_2.active = ws2  # 별도제출자료 정리 엑셀

    for ws1_row_num in range(2, ws1.max_row + 1):
        for ws2_row_num in range(2, ws2.max_row + 1):
            # * excel_1, excel_2 : BOOKID, 위원명, 질의 로 비교
            if ([str(ws1.cell(row=ws1_row_num, column=col).value).strip()
                 for col in [4, 3, 6]] ==
                [str(ws2.cell(row=ws2_row_num, column=col).value).strip()
                 for col in [4, 7, 8]] and
                    ws2.cell(row=ws2_row_num, column=7).value is not None):
                ws2.cell(row=ws2_row_num, column=5, value=ws1.cell(
                    row=ws1_row_num, column=5).value)  # SEQNO

    return excel_2


def insert_filename_data(input_path, file_id):
    """3. BOOKID, SEQNO를 모두 병합 후, 순서대로 FILENAME을 삽입합니다"""
    wb = load_workbook(input_path)
    ws = wb.worksheets[1]
    wb.active = ws

    file_id_length = len(file_id)
    file_id_to_int = int(file_id)

    for ws_row_num in range(2, ws.max_row + 1):
        realfile_name = ws.cell(row=ws_row_num, column=11).value
        if realfile_name is None:
            continue
        _, extension = os.path.splitext(realfile_name)
        upper_extension = extension.upper()
        file_name =\
            f"{str(file_id_to_int).zfill(file_id_length)}{upper_extension}"

        ws.cell(row=ws_row_num, column=6, value=file_name)
        file_id_to_int += 1

    wb.save(input_path)
