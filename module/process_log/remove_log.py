"""
로그 파일을 삭제합니다
"""


import os


from module.process_log.get_log_path import get_log_path


def remove_log():
    """로그 파일을 삭제합니다"""
    log_input = input("로그 파일을 정말 삭제하겠습니까? (Y/N) : ").lower()

    match log_input:
        case 'y':
            if os.path.exists(get_log_path()):
                os.remove(get_log_path())
                print("=> 로그 파일이 정상적으로 삭제되었습니다\n")
            else:
                print("=> 로그 파일이 존재하지 않습니다\n")
        case 'n':
            print("=> 로그 파일 삭제가 취소되었습니다\n")
            return
        case _:
            print("=> y 혹은 n 으로 적어주세요\n")
            remove_log()
