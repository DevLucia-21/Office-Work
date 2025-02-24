# PDF_BM_Start_Num.py (자동 책갈피 생성-시작 번호 설정)

PDF 문서에 **페이지 번호를 책갈피로 자동 추가**하는 Python 스크립트
시작 번호를 입력하면, 시작 번호부터 책갈피를 생성한 후, 새로운 PDF 파일로 저장

## 기능

- PDF 파일에서 기존 목차(TOC)를 제거
- 사용자 정의 시작 번호를 기반으로 각 페이지에 책갈피를 추가
- 수정된 PDF 파일을 지정된 경로에 저장

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
python PDF_BM_Start_Num.py
```

### 2. PDF 파일 경로 입력

프로그램이 실행되면 **입력할 PDF 파일 경로**와 **출력할 PDF 파일 경로**를 입력

```
PDF 파일 경로를 입력하세요: C:\Users\yeeun\Desktop\input.pdf
출력 PDF 파일 경로를 입력하세요: C:\Users\yeeun\Desktop\output.pdf
```

### 3. 시작 번호 입력

시작할 책갈피 번호를 입력

```python
시작 번호를 입력하세요:  n
```

### 4. 결과 확인

출력된 PDF 파일을 열어 **책갈피가 정상적으로 추가되었는지 확인**

## 코드 설명

### 핵심 코드

```
import fitz  # PyMuPDF

def add_bookmarks_to_pdf(pdf_path, output_path, start_number):
    # PDF 파일 열기
    pdf_document = fitz.open(pdf_path)

    # 책갈피 추가
    toc = []
    for page_number in range(pdf_document.page_count):
        title = str(start_number + page_number)  # 시작 번호를 기준으로 책갈피 설정
        toc.append([1, title, page_number])  

    # 변경사항 저장
    pdf_document.save(output_path)
    pdf_document.close()
```

## 주의 사항

- PDF 파일 경로는 **절대 경로**로 입력
- 기존 책갈피는 모두 제거
- `pymupdf`가 설치되지 않았다면 실행되지 않음