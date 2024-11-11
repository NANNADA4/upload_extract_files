"""
두 엑셀 파일을 비교합니다.
"""


from openpyxl import Workbook


def add_pdf_answer(excel_1: Workbook, excel_2: Workbook) -> Workbook:
    """두 엑셀 파일을 비교하여 값이 같을 경우 PDF상 답변을 추가합니다"""
    ws1 = excel_1.active  # 1번 스크립트에서 생성한 엑셀파일
    ws2 = excel_2.worksheets[1]
    excel_2.active = ws2  # 별도제출자료 정리 엑셀

    for ws1_row_num in range(2, ws1.max_row + 1):
        for ws2_row_num in range(2, ws2.max_row + 1):
            # * excel_1, excel_2 : 위원회, 피감기관, 위원명, 질의 로 비교
            if ([str(ws1.cell(row=ws1_row_num, column=col).value).strip()
                 for col in [1, 2, 3, 6]] ==
                [str(ws2.cell(row=ws2_row_num, column=col).value).strip()
                 for col in [1, 2, 5, 6]] and
                    ws2.cell(row=ws2_row_num, column=5).value is not None):
                ws2.cell(row=ws2_row_num, column=3, value=ws1.cell(
                    row=ws1_row_num, column=4).value)  # BOOKID
                ws2.cell(row=ws2_row_num, column=4, value=ws1.cell(
                    row=ws1_row_num, column=5).value)  # SEQNO

    return excel_2

    #! 엑셀 파일 생성 후 정렬 후 FILENAME 삽입
