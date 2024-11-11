"""
두 엑셀 파일을 비교합니다.
"""

import os
from collections import OrderedDict
from openpyxl import Workbook


def add_pdf_answer(excel_1: Workbook, excel_2: Workbook) -> Workbook:
    """두 엑셀 파일을 비교하여 값이 같을 경우 PDF상 답변을 추가합니다"""
    total_dic = {}

    ws1 = excel_1.active  # 1번 스크립트에서 생성한 엑셀파일
    ws2 = excel_2.worksheets[1]
    excel_2.active = ws2  # 별도제출자료 정리 엑셀

    for ws1_row_num in range(2, ws1.max_row + 1):
        attach_excel1_info_list = []
        pdf_answer_list = []  # 질의별 별첨파일 리스트

        for ws2_row_num in range(2, ws2.max_row + 1):
            # * excel_1, excel_2 : 위원회, 피감기관, 위원명, 질의 로 비교
            if ([str(ws1.cell(row=ws1_row_num, column=col).value).strip() for col in [1, 2, 3, 6]] ==  # pylint: disable=C0301
                [str(ws2.cell(row=ws2_row_num, column=col).value).strip() for col in [1, 2, 4, 5]] and  # pylint: disable=C0301
                    ws2.cell(row=ws2_row_num, column=5).value is not None):
                pdf_answer_list.append(
                    ws2.cell(row=ws2_row_num, column=5).value)

        if len(pdf_answer_list) > 0:
            attach_excel1_info_list.append(
                [ws1.cell(row=ws1_row_num, column=1).value,  # 위원회
                 ws1.cell(row=ws1_row_num, column=2).value,  # 피감기관
                 ws1.cell(row=ws1_row_num, column=3).value,  # 위원명
                 ws1.cell(row=ws1_row_num, column=4).value,  # BOOKID
                 ws1.cell(row=ws1_row_num, column=5).value,  # SEQNO
                 ws1.cell(row=ws1_row_num, column=6).value,  # 질의
                 ws1.cell(row=ws1_row_num, column=9).value,  # 파일명
                 pdf_answer_list])  # PDF상 답변

            total_dic.update(
                {ws1_row_num + 1: attach_excel1_info_list})

    # 역순으로 dic 변환 후 엑셀 삽입
    sorted_keys = sorted(total_dic.keys(), reverse=True)
    final_dict = OrderedDict((key, total_dic[key]) for key in sorted_keys)

    #! dictionary 디버깅용. PyInstaller사용시 오류 발생함. 사용X
    # with open('./log/data.txt', 'w', encoding='UTF-8') as file:
    #     for key, value in final_dict.items():
    #         file.write(f"{key}: {value}\n")

    for key, value in final_dict.items():
        ws1.insert_rows(key, len(value[0][7]))
        for cnt in range(len(value[0][7])):
            ws1.cell(row=key + cnt, column=1, value=value[0][0])
            ws1.cell(row=key + cnt, column=2, value=value[0][1])
            ws1.cell(row=key + cnt, column=3, value=value[0][2])
            ws1.cell(row=key + cnt, column=4, value=value[0][3])
            ws1.cell(row=key + cnt, column=5, value=value[0][4])
            ws1.cell(row=key + cnt, column=6, value=value[0][5])
            ws1.cell(row=key + cnt, column=9, value=value[0][6])
            ws1.cell(row=key + cnt, column=7, value=value[0][7][cnt])
    # remove_rows(ws1)
        ws1.delete_rows(key - 1)

    return excel_1


def add_attach_list(wb1: Workbook, wb2: Workbook, file_id: str) -> Workbook:
    """PDF상 답변을 병합 뒤, 별첨파일이름과 FILE_NAME을 입력합니다"""
    ws1 = wb1.active
    ws2 = wb2.worksheets[1]  # 2번째 sheet 실행
    wb2.active = ws2

    attach_total_dic = {}

    for ws1_row_num in range(2, ws1.max_row + 1):
        attach_excel1_info_list = []
        file_name_list = []  # 질의별 별첨파일 리스트
        for ws2_row_num in range(2, ws2.max_row + 1):
            # * 위원회, 피감기관, 위원명, 질의, PDF상 답변 으로 비교
            if ([str(ws1.cell(row=ws1_row_num, column=col).value).strip() for col in [1, 2, 3, 6, 7]] ==  # pylint: disable=C0301
                [str(ws2.cell(row=ws2_row_num, column=col).value).strip() for col in [1, 2, 4, 5, 6]] and  # pylint: disable=C0301
                    ws2.cell(row=ws2_row_num, column=9).value is not None):
                file_name_list.append(
                    ws2.cell(row=ws2_row_num, column=9).value)

        if len(file_name_list) > 0:
            attach_excel1_info_list.append(
                [ws1.cell(row=ws1_row_num, column=1).value,  # 위원회
                 ws1.cell(row=ws1_row_num, column=2).value,  # 피감기관
                 ws1.cell(row=ws1_row_num, column=3).value,  # 위원명
                 ws1.cell(row=ws1_row_num, column=4).value,  # BOOKID
                 ws1.cell(row=ws1_row_num, column=5).value,  # SEQNO
                 ws1.cell(row=ws1_row_num, column=6).value,  # 질의
                 ws1.cell(row=ws1_row_num, column=7).value,  # PDF상 답변
                 ws1.cell(row=ws1_row_num, column=9).value,  # 파일명
                 file_name_list])  # 별첨파일명

            attach_total_dic.update({ws1_row_num + 1: attach_excel1_info_list})

    sorted_keys = sorted(attach_total_dic.keys(), reverse=True)
    final_attach_dict = OrderedDict(
        (key, attach_total_dic[key]) for key in sorted_keys)

    #! dictionary 디버깅용. PyInstaller사용시 오류 발생함. 사용X
    # with open('./log/data2.txt', 'w', encoding='UTF-8') as file:
    #     for key, value in final_attach_dict.items():
    #         file.write(f"{key}: {value}\n")

    for key, value in final_attach_dict.items():
        ws1.insert_rows(key, len(value[0][8]))
        for cnt in range(len(value[0][8])):
            ws1.cell(row=key + cnt, column=1, value=value[0][0])  # 위원회
            ws1.cell(row=key + cnt, column=2, value=value[0][1])  # 피감기관
            ws1.cell(row=key + cnt, column=3, value=value[0][2])  # 위원명
            ws1.cell(row=key + cnt, column=4, value=value[0][3])  # BOOK_ID
            ws1.cell(row=key + cnt, column=5, value=value[0][4])  # SEQ_NO
            ws1.cell(row=key + cnt, column=6, value=value[0][5])  # 질의
            ws1.cell(row=key + cnt, column=7, value=value[0][6])  # PDF상 답변
            ws1.cell(row=key + cnt, column=9, value=value[0][7])  # 파일명
            ws1.cell(row=key + cnt, column=8, value=value[0][8][cnt])  # 별첨파일
        ws1.delete_rows(key - 1)

    insert_cell_data(ws1, file_id)

    return wb1


def insert_cell_data(ws, file_id):
    """입력받은 file_id를 이용해 FILE_NAME을 만듭니다"""
    file_id_length = len(file_id)
    file_id_to_int = int(file_id)

    for ws_row_num in range(2, ws.max_row + 1):
        realfile_name = ws.cell(row=ws_row_num, column=8).value
        if realfile_name is None:
            continue
        _, extension = os.path.splitext(realfile_name)
        upper_extension = extension.upper()
        file_name =\
            f"{str(file_id_to_int).zfill(file_id_length)}{upper_extension}"

        ws.cell(row=ws_row_num, column=11, value=file_name)
        file_id_to_int += 1


def remove_rows(ws):
    """별첨파일 데이터가 없는 엑셀의 행을 삭제합니다"""
    delete_rows = []

    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=7).value is None:
            delete_rows.append(row)

    for row in reversed(delete_rows):
        ws.delete_rows(row)
