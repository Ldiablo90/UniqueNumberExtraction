import os
import datetime as dt
import string
import pandas as pd
import serialfunc as sfunc
import numpy as np


original_working_directory = os.getcwd()

networkpath = r"\\Desktop-sl150kj\송장파일"

mainfiles = ["슈케이브통합시리얼이벤트_val_1.xlsx","슈케이브통합시리얼이벤트_val_1.xlsx"]

meridiem = "오전" if dt.datetime.now().hour < 14 else "오후" # if 오전 else 오후


mainSheetName = "이벤트 종합" # 시리얼번호 종합시트 이름
useColumns = ["시리얼번호","상품주문번호","상품명","옵션정보","수량","수취인명","수취인연락처1","기본배송지","상세배송지"]
needColumns = ["상품주문번호","상품명","옵션정보","수량","수취인명","수취인연락처1","기본배송지","상세배송지"]
randomNumber = [i for i in string.ascii_letters] + [str(i) for i in range(10)] # 영문대소문자 + 숫자열

files = os.listdir(networkpath)

if len(files) > 0:
    print("공유파일에 접속하였습니다.")
    todaylastfile = list(filter(sfunc.todayFileFilter, files))[-1]
    dataFilePath = r"%s\%s"%(networkpath,todaylastfile)
    if os.path.isfile(dataFilePath):
        print("오늘의 정보를 찾았습니다.\n%s"%todaylastfile)
        productType = True
        for serialFile in mainfiles:
            dataSheetName = "{}-{}-{}".format(dt.date.today(),meridiem, "프" if productType else "베")  # 현 시간 최신데이터 시트 이름
            mainFilePath = r"%s\%s"%(original_working_directory,serialFile)
            originPd, loadData = sfunc.mainFileok(mainFilePath,mainSheetName) if os.path.isfile(mainFilePath) else sfunc.mainFileno(useColumns) # 메인파일이 있는지 확인 후 파일 읽기
            originList = originPd["시리얼번호"].values # 시리얼번호 모두 가져오기
            lastNum = len(originList)+1 # 시리얼번호 길이
            dataPd = sfunc.todayFileok(originList,lastNum, dataFilePath, productType, needColumns, useColumns) if os.path.isfile(dataFilePath) else sfunc.todayFileno(useColumns)
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
            productType=False
else:
    print("접속에 실패하였습니다.\n수동으로 작업해주세요.")
os.system("pause")
