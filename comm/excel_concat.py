import pandas as pd
import os
import xlrd
def openExcel(path, sheet):
    excel = pd.ExcelFile(path)
    columns = excel.parse(sheet).columns
    # ~ 번호 컬럼들만 텍스트로 변경하지 않으면 자동으로 지수승으로 정수형으로 바껴서 문자열 형으로 변환하는 컨버터 추가
    converters = {}
    # 컬럼명 중 "번호" 라는 문구가 들어가는 컬럼만 숫자가 지수표현으로 바뀌지 않도록 텍스트 변경
    # str 타입으로 읽어야 숫자가 반올림되거나 엑셀에서 자동으로 포매팅해서 이상하게 읽히는 상황이 발생안함
    for col in columns:
        if "번호" in col:
            converters[col] = str
    excel = excel.parse(sheet, converters=converters)
    return excel

# 다운받은 엑셀 파일에 데이터 누락이 없는지 데이터 행 수 반환
def fileRowCheck(path):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_index(0)
    # 컬럼명을 제외한 데이터 행 수
    if "전체주문" in path:
        return ws.nrows - 2
    return ws.nrows - 1

def initExcel(excel, brand):
    # 브랜드명 컬럼 추가
    excel.insert(0, '브랜드명', brand)
    return excel

def concatExcel(main, sub, brand):
    # 브랜드명 컬럼 추가
    sub.insert(0, '브랜드명', brand)
    main = pd.concat([main, sub], ignore_index=True)
    return main

# src: 다운로드 받은 파일들이 존재하는 경로, target: 병합된 파일 생성 경로, line: 병합 시 생략 할 행 수 -> 반품완료, 취소완료: 0, 발주발송:1
# adminLog: 관리자용 텍스트 로거, userLog: 현업용 재다운로드 확인용 로거
def getResultFile(src, target, line, adminLog, userLog):
    flag = True
    emp = None
    result = None
    sheet = None
    # 해당 업무 엑셀 파일의 시트명 찾기, 첫번째 시트로 설정 (반품관리, 취소관리, 발주발송)
    for fname in os.listdir(src):
        fname = r"{}\{}".format(src, fname)
        try:
            workbook = xlrd.open_workbook(fname)
        except Exception as e:
            adminLog.error("{} 파일을 열 수 없어 시트명을 읽어오지 못함! 파일 확인 요망 | {}".format(fname, e))
            continue
        
        for sname in workbook.sheets():
            sheet = sname.name
            break
        break
    
    # 엑셀 파일에서 데이터 읽어와 병합
    for fname in os.listdir(src):
        # 파일명 기준 _ 앞이 브랜드명
        try:
            brand = fname.split('_')
        except Exception:
            adminLog.error("다운 받은 파일명에 _가 존재하지 않음. {} 파일 확인 요망!".format(fname))
            continue

        # 오류 발생 시 로그 기록에 사용할 fname을 파일경로로 변경하기전 이름 저장    
        temp = fname    

        fname = r"{}\{}".format(src, fname)
        try:
            emp = openExcel(fname, sheet)
        except Exception as e:
            adminLog.error("{} 파일의 엑셀 객체를 얻어 오지 못했습니다. 파일 확인 요망! | {}".format(fname, e))
            userLog.info(temp)
            continue

        # 최초의 한번, 첫번째 파일로만 초기화 해주고 나머지 파일들은 그 뒤에 붙이는 방식
        if flag:
            try:
                result = initExcel(emp, brand[0])
            except Exception as e:
                adminLog.error("엑셀 객체 초기화 실패, {} 파일 확인 요망! | {}".format(fname, e))
                userLog.info(temp)
            flag = False
        else:
            emp = emp[line:]
            try:
                result = concatExcel(result, emp, brand[0])
            except Exception as e:
                adminLog.error("엑셀 객체 병합 실패, {} 파일 확인 요망! | {}".format(fname, e))
                userLog.info(temp)



    # 일매출정리 발주발송처럼 위에 2줄 제거하고 합치는 파일의 경우
    # 병합을 하게 되면 컬럼명들은 자동으로 생략되고 밑에 데이터부터 병합되므로 맨 마지막에 맨 윗줄 불필요한 행만 없애주는 작업
    if line == 1:
        # 첫번째 행에 다운받았을때 있는 컬럼명이 아닌 그 아래칸의 실제 사용되는 컬럼명이 있는 첫번째 행을 컬럼명으로 변경
        result.columns = result.loc[0]
        # 실제 사용되는 컬럼명이 있는 행 다음부터 실제 데이터들이므로 1번째 줄은 컬럼명이 되었으므로 2번째 줄 부터 시작하도록 설정
        result = result[1:]

    # 병합된 결과 파일에 쓰기
    try:
        with pd.ExcelWriter(target) as writer:
            result.to_excel(writer, index=False)
        
    except Exception as e:
        adminLog.error("병합 엑셀 생성 실패!!! | {}".format(e))

# path = r"C:\Users\Administrator\Desktop\naver\naverdaily\ddp\2020\09\08\발주발송(발송처리일)"
# target = r'C:\Users\Administrator\Desktop\naver\naverdaily\ddp\2020\09\08\발주발송(발송처리일).xlsx'
# adminLog = api.getAdminLogger(r'C:\Users\Administrator\Desktop\naver\naverdaily\ddp\2020\09\08\log\20200908.log', "naver_daily")
# userLog = api.getUserLogger(r'C:\Users\Administrator\Desktop\naver\naverdaily\ddp\2020\09\08\log\20200908_보고파일.csv', "naver_daily_fail")
# # 발주발송처럼 2줄 지워야되면 1, 나머지는 0
# getResultFile(path, target, 1, adminLog, userLog)