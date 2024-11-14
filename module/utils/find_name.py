"""
파일명에서 위원회명과 피감기관명을 가져옵니다.
"""


def extract_cmt(filename) -> str:
    """파일명에서 위원회 이름 추출"""
    first_underscore_index = filename.find('_')
    second_underscore_index = filename.find(
        '_', first_underscore_index + 1)
    if first_underscore_index != -1 and second_underscore_index != -1:
        cmt = filename[first_underscore_index +
                       1:second_underscore_index]
    else:
        cmt = ""

    return cmt


def extract_org(filename) -> str:
    """파일명에서 피감기관 이름 추출"""
    stack = []
    end = None
    org = ""

    for i in range(len(filename) - 1, -1, -1):
        char = filename[i]

        if char == ')':
            stack.append(i)
            if len(stack) == 1:
                end = i
        elif char == '(':
            stack.pop()
            if len(stack) == 0:
                org = filename[i + 1:end]
                break

    return org if org else ""
