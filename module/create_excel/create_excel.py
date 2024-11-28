"""
입력받은 폴더 경로로부터 PDF파일을 입력받고, 엑셀 파일을 생성합니다.
"""

import os
from natsort import natsorted


from module.utils.find_name import extract_cmt, extract_org
from module.utils.load_excel import load_excel
from module.utils.extract_bookmark import extract_bookmark


def create_excel(input_path, excel_path):
    """엑셀 파일을 생성합니다"""
    wb = load_excel(excel_path)
    ws = wb.active

    for root, _, files in os.walk(input_path):
        for file in natsorted(files):
            cmt = extract_cmt(file)
            org = extract_org(file)

            tmp = 1
            last_row = ws.max_row

            for item in extract_bookmark(os.path.join(root, file)):
                if len(item) > 1 and item['level'] == 3:
                    ws.cell(row=last_row + tmp, column=1, value=cmt)
                    ws.cell(row=last_row + tmp, column=2, value=org)
                    ws.cell(row=last_row + tmp, column=3,
                            value=item['parent']['title'])
                    ws.cell(row=last_row + tmp,
                            column=5, value=item['SeqNo'])
                    ws.cell(row=last_row + tmp,
                            column=6, value=item['title'])
                    filename, _ = os.path.splitext(file)
                    ws.cell(row=last_row + tmp, column=9, value=filename)
                    tmp += 1
    wb.save(excel_path)
