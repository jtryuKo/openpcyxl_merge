import pandas
import os
import openpyxl

#1.  폴더내 파일을 검색하기
dir = 'C:/Users/ryuje/Desktop/PyProject/blog6' #디렉토리 위치
files = os.listdir(dir) #폴더에 있는 파일 정보를 가져오기

dataframeFile = pandas.DataFrame(index=range(0, 0), columns=['파일명', '이름', '확장자', '위치정보']) # 파일의 정보를 넣는 데이터 프레임 생성
def file_search(dir, dataframeFile):
    files = os.listdir(dir)
    for file in files:
        fullname_file = os.path.join(dir, file)
        fullname_file = fullname_file.replace("\\", "/")
        if os.path.isdir(fullname_file):
            dataframeFile = file_search(fullname_file, dataframeFile)  # 재귀함수 호출
        else:
            name, ext = os.path.splitext(file)
            dic_file = pandas.DataFrame({'파일명': file, '이름': name, '확장자': ext, '위치정보': fullname_file}, index=[0])
            dataframeFile = pandas.concat([dataframeFile, dic_file], ignore_index=True)
    # 데이터프레임 리턴
    return dataframeFile

dataframeFile = file_search(dir, dataframeFile) # 폴더내 파일을 재귀함수 호출 검색

#2. 폴더내 엑셀파일만 데이터 프레임에 남기기
xldataframe = dataframeFile.where(dataframeFile['확장자']=='.xlsx') #.xlsx만 별도로 추출
xldataframe = xldataframe.dropna() #Na 결측치 제거
xldataframe = xldataframe.reset_index() #index 재설정
print(xldataframe) # xldataframe 출력

#3. 파일 생성하기

DirectoryString = dir+'/'
FileString = 'result'

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

for i in range(0, 2):  # Excel_Export 실행해 2개의 파일을 생성
    Excel_Export(DirectoryString, FileString, i)

#4. 각각 다른 데이터 입력하기
#4-1. 모든 파일 합치기
index = 0 #result0 파일에 저장
for i in range(0, len(xldataframe.index)):
     if i==0:
         XlsxData_dataframe = pandas.read_excel(xldataframe.iloc[i].loc['위치정보'], sheet_name="Sheet1")
         total_dataframe = XlsxData_dataframe
     else:
         XlsxData_dataframe = pandas.read_excel(xldataframe.iloc[i].loc['위치정보'], sheet_name="Sheet1")
         total_dataframe = pandas.concat([total_dataframe, XlsxData_dataframe], ignore_index = True)

writer = pandas.ExcelWriter(('%s%s%s.xlsx'%(DirectoryString, FileString, str(index))), engine = 'xlsxwriter') #생성된 데이터 쓰기
total_dataframe.to_excel(writer, sheet_name = 'Sheet1')
writer.save()
print("result0 완료")

#4-2. 두개의 파일만 합치기
index = 1 #result1 파일에 저장
for i in range(0, len(xldataframe.index)-1):
     if i==0:
         XlsxData_dataframe = pandas.read_excel(xldataframe.iloc[i].loc['위치정보'], sheet_name="Sheet1")
         total_dataframe = XlsxData_dataframe
     else:
         XlsxData_dataframe = pandas.read_excel(xldataframe.iloc[i].loc['위치정보'], sheet_name="Sheet1")
         total_dataframe = pandas.concat([total_dataframe, XlsxData_dataframe], ignore_index = True)

writer = pandas.ExcelWriter(('%s%s%s.xlsx'%(DirectoryString, FileString, str(index))), engine = 'xlsxwriter') #생성된 데이터 쓰기
total_dataframe.to_excel(writer, sheet_name = 'Sheet1')
writer.save()
print("result1 완료")
