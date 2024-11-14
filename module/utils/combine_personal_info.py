"""
위원회, 피감기관, 위원명, 파일명을 비교하여 전부 일치하면 개인정보가 존재한다는 표시를 합니다
"""


import openpyxl


def combine_is_exist_personal_info(file_path):
    """위원회, 피감기관, 위원명, 파일명을 비교하여 전부 일치하면 개인정보가 존재한다는 표시를 합니다"""
    workbook = openpyxl.load_workbook(file_path)

    sheet2 = workbook.worksheets[1]
    sheet4 = workbook.worksheets[3]

    for row2 in range(2, sheet2.max_row + 1):
        val_a2 = sheet2.cell(row=row2, column=1).value
        val_b2 = sheet2.cell(row=row2, column=2).value
        val_e2 = sheet2.cell(row=row2, column=5).value
        val_i2 = sheet2.cell(row=row2, column=9).value

        for row4 in range(2, sheet4.max_row + 1):
            val_a4 = sheet4.cell(row=row4, column=1).value
            val_b4 = sheet4.cell(row=row4, column=2).value
            val_c4 = sheet4.cell(row=row4, column=3).value
            val_d4 = sheet4.cell(row=row4, column=4).value

            if (val_a2 == val_a4 and val_b2 == val_b4 and val_e2 == val_c4 and val_i2 == val_d4):
                sheet2.cell(row=row2, column=16).value = 'O'

    workbook.save(file_path)
