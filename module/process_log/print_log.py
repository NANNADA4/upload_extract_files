"""
로그 파일을 출력합니다
"""


import os


from module.process_log.get_log_path import get_log_path


def print_log():
    """log 파일 내용을 CMD창에 PRINT합니다"""
    if os.path.exists(get_log_path()):
        with open(get_log_path(), 'r', encoding='UTF-8') as file:
            lines = file.readlines()
        print("="*30)
        for line in lines:
            print(line)
        print("="*30)
    else:
        print("=> 로그 파일이 존재하지 않습니다\n")
