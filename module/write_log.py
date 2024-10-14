"""
변환에 실패한 파일을 로그로 남깁니다
"""


def write_log(log_path, file_path):
    """변환에 실패한 파일을 로그로 남깁니다"""
    with open(log_path, 'a', encoding='UTF-8') as file:
        file.write(file_path)
