# 문서 작성 자동화 & PDF 활용



> 직접 파이썬을 가지고 문서를 작성하는 것은 python-docx 모듈을 통해 가능하지만, 선호되지 않음.
>
> 데이터를 가지고 일정한 형식의 여러 문서를 만들 때는 mailmerge 활용을 권장





### 1. 문서 출력

```python
os.startfile(출력파일경로, "print")
```

- print 함수는 기본 프린터로 문서를 출력하므로 기본 프린터 설정을 확인해야 함.





### 2. mail-merge를 이용해 문서 생성하기

> template가 있을 때 엑셀의 데이터를 대입하여 개인별 문서를 생성함

- template : 문서 형식

- merge field : 데이터값에 따라 달라지는 부분

  - 삽입 - 빠른 문서 요소 - 필드 - 필드 이름 - MergeField -> 고유 필드 이름 지정

  > 해당 문서를 namecard.docx 라고 하자

  ```
  이름 : <<given_name>> <<family_name>>
  나이 : <<age>>
  ```



> 활용

##### 1. template을 MailMerge 객체로

```python
from mailmerge import MailMerge


template = "./namecard.docx"
document = MailMerge(template)
```



##### 2. merge method를 통해 merge field에 data값 넘겨주기

````python
# 문서에 존재하는 merge field 확인
document.get_merge_fields()

# merge() 에 merge field 값 넘겨주기
document.merge(
	given_name='SuHyeon',
    family_name='Lee',
    age='26'
)

# namecard_output.docx에 문서가 저장됨
document.write('namecard_output.docx')
````



##### 3. table과 같이 한 문서에 여러 줄 저장

```python
# merge_row 사용
people = [
    {
        'given_name' : 'SuHyeon',
        'family_name' : 'Lee',
        'age' : '26'
    },
    {
        'given_name' : 'SuBin',
        'family_name' : 'Lee',
        'age' : '31'
    }
]
document.merge_row('given_name', people)
document.write('name_table.docx')
```



##### 4. 같은 형식의 문서를 여러 번 저장

```python
infos = [('SuHyeon', 'Lee', '26'), ('Subin', 'Lee', '31')]

for info in infos:
    document.merge(given_name=info[0], family_name=info[1], age=info[2])
    document.write(f'namecard_{given_name}.docx')
```





### 3. 워드를 PDF로 저장하기

> pywin32, SavdAs(경로, FileFormat=17)

```python
import os
from win32com.client import Dispatch

# 워드 객체 모델의 최상위 객체 Application
# COM객체를 생성하는 CoClass Word.Application
wordapp = Dispatch("Word.Application")
#os.getcwd() 현재 위치의 directory
fpath = os.path.join(os.getcwd(), "test-output-table.docx")
myDoc = wordapp.Documents.Open(FileName=fpath)

# os.path.join() ===> 경로+파일명 합쳐줌
# 분할된 경로를 하나로 정리
pdf_path = os.path.join(os.getcwd(), "test_saved.pdf")
```



###### mail-merge를 이용해 만들었던 문서를 pdf로 저장하기

```python
from mailmerge import MailMerge
import os
from win32com.client import Dispatch


template = "./namecard.docx"
document = MailMerge(template)

# 문서에 존재하는 merge field 확인
document.get_merge_fields()

# merge() 에 merge field 값 넘겨주기
document.merge(
	given_name='SuHyeon',
    family_name='Lee',
    age='26'
)

# namecard_output.docx에 문서가 저장됨
document.write('namecard_output.docx')
```





### 4. 다양한 문서를 pdf로 변환

###### 웹 페이지를 PDF로 변환

````python
# pdfkit 모듈 사용
# wkhtmltopdf ==> webkit html to pdf 설치 필요
import pdfkit


options = {'quiet': ''} # wkhtmltopdf 실행 결과가 뜨지 않도록
config = pdfkit.configuration(wkhtmltopdf=r'wkhtmltopdf경로')

pdfkit.from_url('http://naver.com', 'naver.pdf', options=options, configuration=config)
````



###### excel, ppt를 PDF로 변환

- word

```python
# win32로 읽은 후 PDF로 저장. SaveAs() method에 FileFormat=17
# 코드는 3. 워드를 PDF로 저장하기 참고
import os
from win32com.client import Dispatch

# 워드 객체 모델의 최상위 객체 Application
# COM객체를 생성하는 CoClass Word.Application
wordapp = Dispatch("Word.Application")
#os.getcwd() 현재 위치의 directory
fpath = os.path.join(os.getcwd(), "test-output-table.docx")
myDoc = wordapp.Documents.Open(FileName=fpath)

# os.path.join() ===> 경로+파일명 합쳐줌
# 분할된 경로를 하나로 정리
pdf_path = os.path.join(os.getcwd(), "test_saved.pdf")
```

- excel

```python
# win32로 읽은 후 PDF로 저장. ExportAsFixedFormat() method 사용
import os
from win32com.client import Dispatch


excelapp = Dispatch("Excel.Application")
excelapp.Visible = False

# 읽기
fpath = os.path.join(os.getcwd(), "excel/test.xlsx")
wb = excelapp.Workbooks.Open(fpath)

# pdf변경 저장
fpath = os.path.join(os.getcwd(), "excel/test.pdf") # 현재위치+폴더 --> 절대경로
wb.ExportAsFixedFormat(0, fpath) # xlTypePDF : 0

wb.Close()
excelapp.Quit()
```

- PPT

```python
# win32로 읽은 후 PDF로 저장. SaveAs method 두 번째 인자로 32 전달
import os
from win32com.client import Dispatch


pptapp = Dispatch("PowerPoint.Application")
pptapp.Visible = False

fpath = os.path.join(os.getcwd(), "ppt/test.pptx")
ppt = pptapp.Presentations.Open(fpath)

pdf_path = os.path.join(os.getcwd(), "test_saved.pdf")
ppt.SaveAs(fpath, 32)

ppt.Close()
pptapp.Quit()
```





### 5. 여러 PDF 합치기

> pyPDF2의 pdfFileMerger 객체에 PdfFileReader 객체를 추가(append)하여 출력

```python
from pyPDF2 import PdfFileMerger, PdfFileReader


filenames = glob.glob('data/*.pdf')

merger = PdfFileMerger()
for filename in filenames:
    merger.append(PdfFileReader(open(filename, 'rb')))

merger.write("total_PDF.pdf")
```





### 6. PDF 분할

> PdfFileReader로 페이지를 읽은 후 getPage(i)로 객체에 접근
>
> 합칠 페이지만 PdfFileWriter객체의 addPage() method로 다시 합치기 

```python
from PyPDF2 import PdfFileWriter, PdfFileReader


inputpdf = PdfFileReader(open("data/sample_book.pdf", "rb"))

parts = [(0, 5), (5, 10)] # 5페이지씩
for idx, (start, end) in enmerate(parts):
    end = min(end, inputpdf.numPages)
    output = PdfFileWriter() # 빈 객체
    # 전체 page에서 해당하는 쪽수만 합쳐서 새로운 객체로
    for i in range(start, end):
        output.addPage(inputpdf.getPage(i))
    
    # newpdf_0, newpdf_1 이라는 이름으로 저장하기
    with open(f"newpdf_{idx}.pdf", "wb") as outputStream:
        output.write(outputStream)
```





### 7. PDF 페이지 제거

> pikepdf module 사용

<mark>페이지 번호가 큰 쪽부터 제거</mark> ==> 앞 쪽 먼저 제거할 시, 페이지 번호 당겨져서! ! !

```python
import pikepdf


with pikepdf.open('data/sample_book.pdf') as pdf:
    num_pages = len(pdf.pages)
    pages_to_delete = [2, 4, 6, num_pages-1]
    # 페이지 큰 번호부터(역순으로) 제거! ! !
    for pg in sorted(pages_to_delete, reverse=True):
        del pdf.pages[pg]
    pdf.save('data/sample_book_deleted_pages.pdf')
```

