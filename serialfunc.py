import random as rd
import string
import openpyxl
import pandas as pd
import numpy as np
import datetime as dt

randomNumber = [i for i in string.ascii_letters] + [str(i) for i in range(10)] # 영문대소문자 + 숫자열


# 시리얼번호 생성 함수
def makeSerialNumber(i):
    # 섞인문자열 중 6개 를 뽑고 순번을 매긴다.
    temp = ''.join(rd.choices(randomNumber,k=6))+ "-" + "000{}".format(i)[-3:]
    return temp

# 메인파일 검색 성공 함수
def mainFileok(mainFilePath, mainSheetName):
    print(mainFilePath+" 파일 있음")
    return pd.read_excel(mainFilePath,sheet_name=mainSheetName), openpyxl.load_workbook(mainFilePath)

# 메인파일 검색 실패 후 함수
def mainFileno(useColumns):
    print("파일 없음 새로생성")
    return pd.DataFrame(columns=useColumns), openpyxl.Workbook()

# 현시점 최신파일 검색 성공 후 함수
def todayFileok(originList, lastNum, dataFilePath, productType, needColumns, useColumns):
    nowDataTemp = [] # 임시저장공간
    tmepNum = lastNum
    dataPd = pd.read_excel(dataFilePath) # 현 시간 최신데이터 파일 가져오기
    if(productType):
        print("프리미엄 검사...")
        dataPd = dataPd[dataPd["상품명"].str.contains("프리미엄")].sort_values(by=["수량"],ascending=False) # 배송파일 프리미엄 구분 / 수량으로 정렬
    else:
        print("일반형 검사...")
        dataPd = dataPd[dataPd["상품명"].str.contains("투명 와이드")].sort_values(by=["수량"],ascending=False) # 배송파일 일반형 구분 / 수량으로 정렬
    
    dataList = (dataPd[needColumns].values) # 배송파일 필요한 정보만 남기기
    for orderNumber in dataList:
        rd.shuffle(randomNumber) # 문자열 순서 섞기
        serialNumber = makeSerialNumber(tmepNum) # 시리얼넘버 생성
        while True:
            if any(np.isin(originList,[serialNumber])): # 시리얼넘버 중복확인
                serialNumber = makeSerialNumber(tmepNum) # 중복이면 다시 시리얼번호를 생성하고 while문으로 돌아간다
            else:
                tmepNum +=1 # 다음순서번호
                temp = [str(x) for x in np.insert(orderNumber,0,serialNumber)] # 리스트화 및 필요한정보 맨 앞에 시리얼넘버 입력
                nowDataTemp.append(temp) # 임시저장
                break
    todayPd = pd.DataFrame(nowDataTemp, columns=useColumns) # 백터정보를 데이터프레임화
    return todayPd

# 현시점 최신파일 검색 실패 후 함수
def todayFileno():
    return pd.DataFrame()

def todayFileFilter(x:str):
    today = "%s"%dt.date.today()
    return today in x