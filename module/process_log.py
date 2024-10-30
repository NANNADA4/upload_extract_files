"""
변환에 실패한 파일을 로그로 남깁니다
"""

import os
import sys


def create_log(file_path):
    """변환에 실패한 파일을 로그로 남깁니다"""
    with open(_get_log_path(), 'a', encoding='UTF-8') as file:
        file.write(f"{file_path}\n")


def print_log():
    """log 파일 내용을 CMD창에 PRINT합니다"""
    if os.path.exists(_get_log_path()):
        with open(_get_log_path(), 'r', encoding='UTF-8') as file:
            lines = file.readlines()
        print("="*30)
        for line in lines:
            print(line)
        print("="*30)
    else:
        print("=> 로그 파일이 존재하지 않습니다\n")


def _get_log_path():
    """패키지, 스크립트에 따른 다른 log PATH를 반환합니다"""
    if hasattr(sys, 'frozen'):
        tmp_path = os.path.expanduser('~')
        os.makedirs(f'{tmp_path}\\Desktop\\log_attach_upload', exist_ok=True)
        return f'{tmp_path}\\Desktop\\log_attach_upload\\attach_upload.log'
    os.makedirs('./log', exist_ok=True)
    return './log/attach_upload.log'


def remove_log():
    """로그 파일을 삭제합니다"""
    log_input = input("로그 파일을 정말 삭제하겠습니까? (Y/N) : ").lower()

    match log_input:
        case 'y':
            if os.path.exists(_get_log_path()):
                os.remove(_get_log_path())
                print("=> 로그 파일이 정상적으로 삭제되었습니다\n")
            else:
                print("=> 로그 파일이 존재하지 않습니다\n")
        case 'n':
            print("=> 로그 파일 삭제가 취소되었습니다\n")
            return
        case _:
            print("=> y 혹은 n 으로 적어주세요\n")
            remove_log()
