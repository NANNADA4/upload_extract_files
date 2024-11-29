"""
main 함수
"""


from module.__process__.process_folder import process_folder
from module.__process__.process_log import print_log, remove_log
from module.utils.manual import describe_program


def main():
    """main 함수. PDF의 4단계에서 별도제출자료 리스트를 추출합니다"""
    print("="*60)
    print("\n>>>>>>별도제출자료 업로드 파일 생성기 (폴더 생성 및 파일 변경)<<<<<<\n")
    print("***'실제 파일명', 'FILE_NAME', 'FILE_PATH' 열이 존재하는지 확인하세요***\n")
    print("="*60)

    print("[1] SeqNO가 포함된 엑셀 파일 생성")
    print("[2] SeqNO 병합하기")
    print("[3] FILE_NAME 병합하기 (!!엑셀 파일 정렬 후 사용!!)")
    print("[4] BOOK_ID를 토대로 파일명 변경 및 폴더 생성")
    print("[5] 로그파일로 누락파일 엑셀 생성하기")
    print("[6] 개인정보추출 리스트와 병합하기 (P열에 저장됨)")
    print("[7] 도움말")
    print("[8] 로그파일 출력")
    print("[9] 로그파일 삭제")
    print("[0] 프로그램 종료")
    input_num = input("=> ")

    match input_num:
        case '1' | '2' | '3' | '4' | '5' | '6':
            process_folder(input_num)
        case '7':
            describe_program()
        case '8':
            print_log()
        case '9':
            remove_log()
        case '0':
            return 0
        case _:
            print("\n====올바른 숫자를 입력해주세요====\n")

    return main()


if __name__ == "__main__":
    main()
