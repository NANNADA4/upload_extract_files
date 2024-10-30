"""
입력받은 폴더 경로로부터 PDF파일을 입력받고, 엑셀 파일을 생성합니다.
"""

import os
from natsort import natsorted
from openpyxl import Workbook, load_workbook


from module.create_excel.find_name import extract_cmt, extract_org
from module.create_excel.load_excel import load_excel


def create_excel(input_path, excel_path):
    """엑셀 파일을 생성합니다"""
    ws = load_excel(excel_path).active

    book_id = input("BOOK_ID 시작번호를 입력해주세요 (오름차순으로 매겨집니다) : ")

    for root, _, files in os.walk(input_path):
        for file in natsorted(files):
            cmt = extract_cmt(file)
            org = extract_org(file)
