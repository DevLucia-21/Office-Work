# PDF_BookMark.py (자동 책갈피 생성기)

PDF 문서에 **페이지 번호를 책갈피로 자동 추가**하는 Python 스크립트

사용자가 입력한 PDF 파일을 읽고, 각 페이지에 해당하는 책갈피를 생성한 후, 새로운 PDF 파일로 저장

## 기능

- PDF의 **모든 페이지에 맞는 책갈피를 자동으로 추가**
- 책갈피 이름은 페이지 번호로 설정 (1, 2, 3, ...)
- 페이지 번호가 1부터 시작하도록 설정
- 기존 책갈피를 제거하고 새로운 책갈피를 추가

## 설치 방법

### 필수 라이브러리 설치

`pymupdf`(PyMuPDF) 라이브러리를 사용. 아래 명령어로 설치:

```
pip install pymupdf
```

또는 Anaconda 환경에서 실행할 경우:

```
conda install -c conda-forge pymupdf
```

## 사용 방법

### 1. 스크립트 실행

```
python PDF_BookMark.py
```

### 2. PDF 파일 경로 입력

프로그램이 실행되면 **입력할 PDF 파일 경로**와 **출력할 PDF 파일 경로**를 입력

```
PDF 파일 경로를 입력하세요: C:\Users\yeeun\Desktop\input.pdf
출력 PDF 파일 경로를 입력하세요: C:\Users\yeeun\Desktop\output.pdf
```

### 3. 결과 확인

출력된 PDF 파일을 열어 **책갈피가 정상적으로 추가되었는지 확인**

## 코드 설명

### 핵심 코드

```
import fitz  # PyMuPDF

def add_bookmarks_to_pdf(pdf_path, output_path):
    pdf_document = fitz.open(pdf_path)
    pdf_document.set_toc([])  # 기존 책갈피 제거

    toc = []
    for page_number in range(pdf_document.page_count):
        title = str(page_number + 1)  # 책갈피 제목 (1부터 시작)
        toc.append([1, title, page_number + 1])  # 책갈피 추가

    pdf_document.set_toc(toc)
    pdf_document.save(output_path)
    pdf_document.close()
```

## 주의 사항

- PDF 파일 경로는 **절대 경로**로 입력
- 기존 책갈피는 모두 제거
- `pymupdf`가 설치되지 않았다면 실행되지 않음