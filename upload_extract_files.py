"""
main 함수
"""


from module.process.process_folder import process_folder
from module.man.manual_program import describe_program
from module.process_log.print_log import print_log
from module.process_log.remove_log import remove_log


def main():
    """main 함수. PDF의 4단계에서 별도제출자료 리스트를 추출합니다"""
    print("="*60)
    print("\n>>>>>>별도제출자료 업로드 파일 생성기 (폴더 생성 및 파일 변경)<<<<<<\n")
    print("***'실제 파일명', 'FILE_NAME', 'FILE_PATH' 열이 존재하는지 확인하세요***\n")
    print("="*60)

    print("[1] BOOK_ID, SeqNO가 포함된 엑셀 파일 생성")
    print("[2] 생성된 엑셀 파일에 질의합치기")
    print("[3] BOOK_ID를 토대로 파일명 변경 및 폴더 생성")
    print("[5] 도움말")
    print("[8] 로그파일 출력")
    print("[9] 로그파일 삭제")
    print("[0] 프로그램 종료")
    input_num = input("=> ")

    match input_num:
        case '1' | '2' | '3':
            process_folder(input_num)
        case '5':
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
