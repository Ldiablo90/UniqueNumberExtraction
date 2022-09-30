import string
import random as rd
from xmlrpc.client import Boolean
import pandas as pd
import numpy as np
import os,os.path
import openpyxl
from datetime import date,datetime
import serialfunc as sfunc

#  True = 프리미엄 / False = 일반형
productType = Boolean(int(input("1 or 0 둘중 한가지 입력해주세요\n1:프리미엄  0:일반형  :  ")))

mainFilePath = "./슈케이브프리미엄시리얼이벤트_val_1.xlsx" if productType else "./슈케이브일반형시리얼이벤트_val_2.xlsx"

meridiem = "오전" if datetime.now().hour < 14 else "오후" # if 오전 else 오후
dataFilePath = "./data/스마트스토어_{}_{}.xlsx".format(meridiem,date.today()) # 현 시간 최신데이터 파일경로 이름
dataSheetName = "{}-{}".format(date.today(),meridiem) # 현 시간 최신데이터 시트 이름

mainSheetName = "이벤트 종합" # 시리얼번호 종합시트 이름
useColumns = ["시리얼번호","상품주문번호","상품명","옵션정보","수량","수취인명","수취인연락처1","기본배송지","상세배송지"]
needColumns = ["상품주문번호","상품명","옵션정보","수량","수취인명","수취인연락처1","기본배송지","상세배송지"]
randomNumber = [i for i in string.ascii_letters] + [str(i) for i in range(10)] # 영문대소문자 + 숫자열

originPd, loadData = sfunc.mainFileok(mainFilePath,mainSheetName) if os.path.isfile(mainFilePath) else sfunc.mainFileno(useColumns) # 메인파일이 있는지 확인 후 파일 읽기

originList = originPd["시리얼번호"].values # 시리얼번호 모두 가져오기
lastNum = len(originList)+1 # 시리얼번호 길이

# 현시간 최신파일이 있는지 확인 후 파일 바꾸기
dataPd = sfunc.todayFileok(originList,lastNum, dataFilePath, productType, needColumns, useColumns) if os.path.isfile(dataFilePath) else sfunc.todayFileno() 

# 최신정보의 데이터 확인하기 
if not dataPd.empty:
    writer = pd.ExcelWriter(mainFilePath,engine="openpyxl")
    for sheetname in loadData.sheetnames:
        if sheetname == "Sheet" : continue # 기본시트면 되돌아가기
        tempcolumn = list([i for i in loadData[sheetname].values][0])
        tempvalues = list([i for i in loadData[sheetname].values][1:])
        tempPd = pd.DataFrame(tempvalues,columns=tempcolumn)
        tempPd.to_excel(writer,sheet_name=sheetname,index=False)
    if any(np.isin(loadData.sheetnames,[dataSheetName])):
        print("최신정보의 시트가 이미 존재합니다.")
    else:
        print("시트가 저장 되었습니다.\n 시트네임 : "+dataSheetName)
        enddataPd = pd.concat([originPd, dataPd])
        enddataPd.to_excel(writer,sheet_name=mainSheetName,index=False) # 메인파일에 시리얼번호 종합시트 다시 작성 
        dataPd.to_excel(writer, sheet_name=dataSheetName, index=False) # 메인파일에 현 시간 최신데이터 작성
    writer.save()
else:
    print("최신데이터가 없습니다.\n 경로"+dataFilePath)

print("파일 종료")
os.system("pause")