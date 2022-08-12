import openpyxl

#예시1 엑셀파일 생성하기
wb = openpyxl.Workbook()  #! Workbook 반드시 대문자로 작성해야만 함
#현재 Active된 시트를 지정
ws = wb.active
#엑셀 파일의 첫 번째 워크 시트이름을 정의, 기본은 'Sheet'
ws.title  = 'Sheet1'
new_filenanme = 'C:/Users/ryuje/Desktop/test.xlsx'
wb.save(new_filenanme)
wb.close()






