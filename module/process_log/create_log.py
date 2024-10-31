"""
로그 파일을 생성합니다
"""


from module.process_log.get_log_path import get_log_path


def create_log(file_path):
    """변환에 실패한 파일을 로그로 남깁니다"""
    with open(get_log_path(), 'a', encoding='UTF-8') as file:
        file.write(f"{file_path}\n")
