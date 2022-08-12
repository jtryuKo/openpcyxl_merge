import openpyxl

#예시2 여러 개의 엑셀파일 생성하기
DirectoryString = 'C:/Users/ryuje/Desktop/'
FileString = 'test'

# Excel_Export 함수
def Excel_Export(DirectoryString, FileString, index) :
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title  = 'Sheet1'
    # 엑셀 파일의 첫 번째 워크 시트이름을 정의, 기본은 'Sheet'
    new_filenanme = DirectoryString + FileString + str(i) + '.xlsx'
    wb.save(new_filenanme)
    wb.close()
    print(new_filenanme)

# Excel_Export 실행
for i in range(0,5) :
    Excel_Export(DirectoryString, FileString, i)

