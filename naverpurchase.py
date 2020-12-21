import multiprocessing
import json
import os
import sys
from comm import naver
from comm import api
from comm import mail
from comm import excel_concat
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# 프로그램 다운 경로
downPath = "C:/Users/Administrator/Desktop/naver/"
# 크롬 다운로드 윈도우 경로
downPath_win = r"C:\Users\Administrator\Desktop\naver"

# cmd 실행 할 때 받는 매개변수 -> 점포명   ddp = 동대문, gimpo = 김포
shop = sys.argv[1]
dirs = ["구매확정일", "결제일"]
fnames = {"구매확정일":"Decision", "결제일":"Pay"}
task = "naverpurchase"
naver.mkdir(downPath, downPath_win, task, shop, dirs)
# mkdir에서 할당된 다운 경로를 현재 파일 변수로 받아오는 작업
downPath = naver.downPath
downPath_win = naver.downPath_win

def main(stores):
    # 사이트 접속, 드라이버 생성, 일감 분장, 로그인
    driver = naver.initProcess(downPath, downPath_win, shop, stores)
    # 메인 비즈니스 로직
    doProcess(driver, stores)

#멀티프로세싱은 여기서만 조작 가능            #######################################################################
if __name__ == '__main__':
    # 멀티 프로세싱 설정 -> 코어 수만큼 활용, 일감은 csv 파일에서 읽기    ... 코어 수를 4 초과하게 되면 홈페이지가 로봇으로 인식해서 막아버리는 이슈 발생
    num_cores = 4#multiprocessing.cpu_count()
    pool = multiprocessing.Pool(num_cores)
    try:
        stores = api.divideWork("{}\\{}".format(naver.nowPath, naver.data["workFileName"][shop]), num_cores)
    except Exception as e:
        naver.adminLog.error("{}: 일감 분리 실패(일감 파일 읽기 불가능)".format(e))

    pool.map(main, stores)
    pool.close()

    day = "{}{}{}".format(api.getYear(), api.getMonth(), api.getDay())
    files = [r"{}\log\{}{}".format(downPath_win, day, "_Report.csv")]
    line = 0
    msg = naver.data["emailText"][task]
    for dir in dirs:
        # 웹메일에서 첨부 메일명에 한글이 포함되면 첨부가 되지 않아서 영어 이름으로 매칭
        filename = r"{}\{}_{}_{}.xlsx".format(downPath_win, day, shop, fnames[dir])
        try:
            excel_concat.getResultFile(r"{}\{}".format(downPath_win, dir), filename, line, naver.adminLog, naver.userLog)
            naver.adminLog.info("{}파일 정상적으로 생성 완료".format(dir))
        except Exception:
            naver.adminLog.error("{}파일 정상적으로 생성 실패".format(dir))
        naver.setDRM(filename)
        isFileExist = False
        try:
            isFileExist = os.path.isfile(filename)
        except Exception as e:
            naver.adminLog.error("네이버 일매출정리 {}파일이 존재하지 않음 | {}".format(filename, e))

        # 파일이 정상적으로 생성되어 존재하면, 결과파일 메일에 첨부 파일명 등록
        if isFileExist:
            files.append(filename)
        else:
        # 파일이 존재하지 않으면 결과파일 메일 내용에 기재
            msg = "<br>◈{}{}파일이 정상적으로 생성되지 못하였습니다!!! <br>".format(msg, fnames[dir])

    # 메일 수신처 설정
    to = ["chosm10@hyundai-ite.com", "cindy@hyundaihmall.com", "move@hyundai-ite.com"]
    to.append(naver.data["email"][shop])

    try:
        mail.sendmail(to, "({}) 네이버 구매확정_{}".format(day, shop), msg, files)
        naver.adminLog.info("네이버 구매확정 메일 정상 발송 완료")
    except Exception as e:
        naver.adminLog.error("네이버 구매확정 메일 정상 발송 실패 | {}".format(e))

    api.taskkill()

####################################################################################################################

# 각 스토어별 일처리 하는 함수
def doProcess(driver, stores):
    # 0번 인덱스에 pid 존재
    pid = stores[0]
    # pid 이후부터 스토어 리스트 존재
    stores = stores[1:]
    # 모든 일감들에 대해서
    for store in stores:
        store = store[0]
        flag = naver.moveStore(pid, driver, store)
        if not flag:
            continue

        # status -> 구매확정
        for status in naver.data["url"][task]:
            # 사이트 이동이 된 후에 첫번째 취소관리로 넘어가야 정상 이동 가능해서 '스마트스토어센터가' 떳는지 확인
            # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["isLogin"])))
            try:
                naver.log(pid, "{} {} 업무 시작".format(store, status))
                naver.cancleAlert(pid, driver)

                naver.delay(3)
                try:
                    detailJob(pid, driver, store, status)
                except Exception as e:
                    naver.adminLog.error("detailJob 에러 발생 | {}".format(e))
                    naver.cancleAlert(pid, driver)
                    continue
            except Exception:
                naver.adminLog.error("{} {}에서 처리 안된 에러 발생".format(store, status))
                naver.cancleAlert(pid, driver)
                continue

#반품,취소관리, 발주발송 별 세부 엑셀 다운 기능
def detailJob(pid, driver, store, status):
    # ["구매확정일", "결제일"]
    jobs = naver.data["statusList"][status]
    for jobname in jobs:
        naver.log(pid, "{} {} 업무 시작".format(store, jobname))
        # 팝업창 제거 하면서 작업 링크로 이동, 한 화면에서 설정만 바뀌지만 사이트를 재접속하지 않으면 크롬에서 파일을 중복해서 다운로드 받을지 물어보면서 다운로드 되지 않기 때문에 강제 이동
        flag = naver.canclePopup(pid, driver, task, store, status)
        if not flag:
            continue
        job = "{} {}".format(store, jobname)
        #홈페이지 이동 후 딜레이를 준 후에 프레임을 변경해야 정상 적용됨!
        naver.delay(3)
        naver.switchFrame(pid, driver)

        naver.delay(3)
        try:
            select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,naver.data[status]["processStatus"]))))
            naver.delay(2)
            select.select_by_visible_text(jobname)
            naver.log(pid, "{} 선택 성공".format(jobname))
        except Exception as e:
            naver.storeExcept(pid, driver, job, "{} 선택 실패 | {}".format(jobname, e))
            continue        
        
        naver.delay(3)
        # 검색 버튼 클릭
        flag = naver.searchData(pid, driver, job, naver.data[status]["search"])
        if not flag:
            continue

        cnt = 0
        check = 0
        checkList = True
        while cnt < 6:
            naver.delay(1)
            # 조회 목록이 0이면 바로 다음작업 진행
            try:
                # 정상적인 상황이면 무조건 인식이 되므로 명시적 지연시간 줘도 그만큼 소요 안됨, 인식 안되면 홈페이지 이상
                obj = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, naver.data["dataCnt"])))
            except Exception:
                naver.storeExcept(pid, driver, job, "조회 목록 인식 실패")
                checkList = False
                break
            # 6초 전에 조회 결과 뜨면 딜레이 멈추고 진행
            if obj.text != "0":
                check = int(obj.text)
                naver.log(pid, "조회목록 {}건 발생".format(obj.text))
                break
            cnt += 1

        if checkList == False:
            continue

        # 다운할 양이 0건이면 다음작업 안하고 진행
        if check == 0:
            naver.log(pid, "{} 다운 할 데이터 없음(0건)".format(job))
            continue

        # 리스트 0번에 뜰때까지 기다림 -> 타임아웃되면 너무 많아서 늦게뜬 경우임 -> 타임아웃이면 넘어감
        try:
            naver.processByXpath(pid, driver, {"xpath":naver.data[status]["checkList"]}, "click", 10)
            naver.log(pid, "조회 항목 로딩 완료")
        except Exception:
            naver.storeExcept(pid, driver, job, "조회 항목 로딩이 지연돼 다운 불가(10초 초과)")
            continue

        # 조회목록이 존재하면 폴더 비우고 엑셀 다운
        naver.delay(2)

        # 작업 폴더 비우기(파일 이름 변경 후 결과폴더로 이동해서 원래 폴더에는 파일이 없어야하는데 혹시라도 안지워진게 있으면 문제가 되서 한번더 폴더 안의 파일들 삭제)
        naver.initDir(pid)
            
        flag = naver.downloadExcel(pid, driver, job, naver.data[status]["download"])
        if not flag:
            continue

        limit = 60
        path = "{}{}".format(downPath, pid)
        isDownOk = naver.downloadWait(pid, driver, job, store, limit, path, ["스마트"])
        # 다운로드 된 파일의 행 수가 스마트스토어 홈페이지의 조회된 데이터 수와 다르면 수기작업 기재
        if isDownOk != check:
            naver.log(pid, "다운받은 파일의 행 수: {}, 홈페이지 조회된 데이터 수: {}로 상이하여 수기작업 필요!".format(isDownOk, check))
            # 작업 폴더 비우기(데이터 누락된 파일이므로 그냥 삭제)
            naver.initDir(pid)
            naver.storeExcept(pid, driver, job, "{} 다운받은 파일이 데이터 건수가 홈페이지와 상이함".format(store))
            continue

        # 몇초 걸려서 기다렸는데도 경로에 파일이 다운되지 못한 경우 한번 더 체크
        isDownOk = naver.downloadCheck(path, ["스마트"])
        if isDownOk == -1:
            naver.storeExcept(pid, driver, job, "{} 파일이 다운로드 되지 못함".format(store))
            continue

        # 여기 오면 파일이 정상 다운된 상황  
        fnameDict = {"구매확정":""}
        fnameDict["구매확정"] = jobname
        naver.moveFile(pid, driver, job, "{}{}".format(downPath, pid), downPath, store, fnameDict)