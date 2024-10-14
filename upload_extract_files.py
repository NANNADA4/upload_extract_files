"""
main 함수
"""

import os


from module.process_folder import process_folder


def main():
    """main 함수. PDF의 4단계에서 별도제출자료 리스트를 추출합니다"""
    print("-"*24)
    print("\n>>>>>>별도제출자료 업로드 파일 생성기 (폴더 생성 및 파일 변경)<<<<<<\n")
    print("***'실제 파일명', 'FILE_NAME', 'FILE_PATH' 열이 존재하는지 확인하세요***")
    print("-"*24)

    input_path = input(
        "입력 폴더 경로를 입력하세요\n*모든 파일이 존재하는 최상위 폴더 경로 입력*\n(종료는 0을 입력) : ")
    if input_path == '0':
        return 0

    output_path = input("폴더 생성 및 이름 변경된 파일이 위치할 폴더 경로를 입력하세요 : ")

    if not os.path.isdir(input_path):
        print("입력 폴더의 경로를 다시 한번 확인하세요")
        return main()

    excel_path = input("엑셀 경로를 입력하세요 : ")
    input_path = os.path.join('\\\\?\\', input_path)

    process_folder(excel_path, input_path, output_path)

    print("\n~~~파일명 변경 및 별도제출자료 업로드 파일 수정이 완료되었습니다.~~~")

    return main()


if __name__ == "__main__":
    main()
