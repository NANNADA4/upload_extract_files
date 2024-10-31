"""
프로그램 실행 환경에 맞춰 로그 PATH를 반환합니다
"""

import os
import sys


def get_log_path():
    """패키지, 스크립트에 따른 다른 log PATH를 반환합니다"""
    if hasattr(sys, 'frozen'):
        tmp_path = os.path.expanduser('~')
        os.makedirs(f'{tmp_path}\\Desktop\\log_attach_upload', exist_ok=True)
        return f'{tmp_path}\\Desktop\\log_attach_upload\\attach_upload.log'
    os.makedirs('./log', exist_ok=True)
    return './log/attach_upload.log'
