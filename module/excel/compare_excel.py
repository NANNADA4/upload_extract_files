"""
두 엑셀 파일을 비교합니다.
"""


from openpyxl import Workbook


def compare_excel(excel_1=Workbook, excel_2=Workbook) -> Workbook:
    """두 엑셀 파일을 비교하여 값이 같을 경우 질의를 추가합니다"""
    cmp1_compare = [1, 2, 3, 6]  # 위원회, 피감기관, 위원명, 질의
    cmp2_compare = [1, 2, 4, 5]  # 위원회, 피감기관, 위원명, 질의

    ws1 = excel_1.active
    ws2 = excel_2.active

    for ws1_row_num in range(ws1_row_num + 1, 2, -1):
        attach_cnt = 0

        cmp1_value = [ws1.cell(
            row=ws1_row_num, column=col).value for col in cmp1_compare]

        attach_list = []
        for ws2_row_num in range(2, ws2.max_row + 1):
            cmp2_value = [ws2.cell(
                row=ws2_row_num, column=col).value for col in cmp2_compare]

            while cmp1_value == cmp2_value:
                attach_cnt += 1

                attach_list.append(ws2.cell(row=ws2_row_num, column=6).value)

                tmp = [
                    ws1.cell(row=ws1_row_num, column=1).value,  # 위원회
                    ws1.cell(row=ws1_row_num, column=2).value,  # 피감기관
                    ws1.cell(row=ws1_row_num, column=3).value,  # 위원명
                    ws1.cell(row=ws1_row_num, column=4).value,  # BOOK_ID
                    ws1.cell(row=ws1_row_num, column=5).value,  # SEQNO
                    ws1.cell(row=ws1_row_num, column=6).value,  # 질의
                    ws2.cell(row=ws2_row_num, column=6).value   # PDF상 답변
                ]

            ws1.insert_rows(ws1_row_num + 1, attach_cnt)

            for cnt in range(0, attach_cnt):
                for col, value in new_ws1_data.items():
                    ws1.cell(row=ws1_row_num + 1 + cnt,
                             column=col, value=value)

    return excel_1
