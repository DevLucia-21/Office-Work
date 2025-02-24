import fitz  # PyMuPDF

def add_bookmarks_to_pdf(pdf_path, output_path, start_number):
    # PDF 파일 열기
    pdf_document = fitz.open(pdf_path)
    
    # 기존 TOC를 초기화하여 제거
    pdf_document.set_toc([])

    # 책갈피 추가
    toc = []
    for page_number in range(pdf_document.page_count):
        title = str(start_number + page_number)  # 시작 번호를 기준으로 책갈피 설정
        toc.append([1, title, page_number])

    # TOC를 문서에 설정
    pdf_document.set_toc(toc)

    # 변경사항 저장
    pdf_document.save(output_path)
    pdf_document.close()

def main():
    # 사용자로부터 PDF 파일 경로 입력받기
    pdf_path = input("PDF 파일 경로를 입력하세요: ")
    output_path = input("출력 PDF 파일 경로를 입력하세요: ")
    
    # 사용자로부터 시작 번호 입력받기
    start_number = int(input("시작 번호를 입력하세요: "))

    add_bookmarks_to_pdf(pdf_path, output_path, start_number)
    print(f"책갈피가 추가된 PDF가 {output_path}에 저장되었습니다.")

if __name__ == "__main__":
    main()
