import multiprocessing
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
if api.delimeter == '/':
    downPath = "/naver/"
    downPath_win = "/naver/"

# cmd 실행 할 때 받는 매개변수 -> 점포명   ddp = 동대문, gimpo = 김포
shop = sys.argv[1]
dirs = ["정산예정일", "정산완료일", "정산기준일"]
fnames = {"정산예정일":"SettleScheduled", "정산완료일":"SettleFinish", "정산기준일":"SettleStandard"}
task = "naversettle"
naver.mkdir(downPath, downPath_win, task, shop, dirs)
# mkdir에서 할당된 다운 경로를 현재 파일 변수로 받아오는 작업
downPath = naver.downPath
downPath_win = naver.downPath_win

# 프로그램 시작
task_name = '네이버_{}_{}'.format(naver.data["task_name"][task], naver.data["shop"][shop])
bot_ip = api.get_ip()
bot_id = naver.data["bot_id"][bot_ip]
data = {'name': task_name, 'botId': bot_id, 'botIp': bot_ip, 'status':'run'}
try:
    naver.log(0, api.post_api(naver.task_status_url, data))
except Exception as e:
    naver.log(0, e)

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

        # status -> scheduled(정산예정일), standard(정산기준일), finish(정산완료일)
        for status in naver.data["url"][task]:
            # 사이트 이동이 된 후에 첫번째 취소관리로 넘어가야 정상 이동 가능해서 '스마트스토어센터가' 떳는지 확인
            # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["isLogin"])))
            try:
                naver.log(pid, "{} {} 업무 시작".format(store, status))
                naver.cancleAlert(pid, driver)

                flag = naver.canclePopup(pid, driver, task, store, status)
                if not flag:
                    continue

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

#정산예정일, 정산완료일, 정산기준일 별 세부 엑셀 다운 기능
def detailJob(pid, driver, store, status):
    task = "{} {}".format(store, naver.data["statusList"][status])
    try:
        naver.switchFrame(pid, driver)
    except Exception:
        naver.log(pid, "프레임 변경 불필요")

    try:
        #일별 정산내역 / 건별 정산내역 버튼 클릭
        naver.processByXpath(pid, driver, {"xpath":naver.data[status]["moveButton"]}, "click", 5)
        naver.log(pid, "정산내역 이동 버튼 클릭 성공")
    except Exception as e:
        # 인식 안될 경우 앞에서 팝업창 때문에 홈페이지 이동이 안된 것이므로 작업 종료하고 다음 작업 진행
        naver.storeExcept(pid, driver, task, "정산내역 이동 버튼 클릭 실패 | {}".format(e))
        return

    naver.delay(3)
    # 정산완료일, 정산기준일만 상태 설정이 있음, 정산예정일은 없음
    text = {"finish":"정산완료일", "standard":"정산기준일"}
    if status == "finish" or status == "standard":
        try:
            select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,naver.data[status]["processStatus"]))))
            naver.log(pid, "{} Select 칸 포커스 성공".format(text[status]))
        except Exception as e:
            naver.storeExcept(pid, driver, task, "{} Select 칸 포커스 실패".format(text[status]))
            return
        naver.delay(2)
        try:
            select.select_by_visible_text(text[status])
            naver.log(pid, "{} 글자 선택 성공".format(text[status]))
        except Exception as e:
            naver.storeExcept(pid, driver, task, "{} 글자 선택 실패".format(text[status]))
            return                

    common = 'react-datepicker__day--0'
    searchDate = {"start": "", "end": ""}
    searchDate["start"], searchDate["end"] = api.getLastDate()
    searchDate["start"] = '{}{}'.format(common, searchDate["start"].split('.')[2])
    searchDate["end"] =  '{}{}'.format(common, searchDate["end"].split('.')[2])
    naver.log(pid, "시작일: {}, 종료일: {} 받아오기 성공".format(searchDate["start"], searchDate["end"]))

    #조회 시작일, 종료일 설정
    idx = 0
    for date in searchDate:
        try:
            objs = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, naver.data[status]["date"])))
            objs[idx].click()
            naver.log(pid, "{} 날짜 칸 클래스로 클릭 성공".format(date))
            idx += 1
        except Exception as e:
            naver.doExcept(pid, store, driver, "{} 날짜 칸 클릭 실패: ".format(date), e)
            return
        
        naver.delay(3)
        try:
            obj = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, naver.data[status]["prev"])))
            obj[0].click()
            naver.log(pid, "전월 버튼 클릭")
        except Exception as e:
            naver.doExcept(pid, store, driver, "전월 버튼 클릭 실패", e)
            return

        naver.delay(3)
        try:
    
            obj = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, searchDate[date])))
            naver.log(pid, "달력 중 주단위 클래스 오브젝트 클래스로 날짜칸 {} 인식 완료".format(searchDate[date]))
        except Exception as e:
            naver.doExcept(pid, store, driver, "달력 중 주단위 클래스 오브젝트 클래스로 날짜칸 {} 인식 불가".format(searchDate[date]), e)
            return

        naver.delay(3)
        try:
            if date == 'end':
                if len(obj) >= 2:
                    obj[1].click()
                else:
                    obj[0].click()
            elif date == 'start':
                obj[0].click()
            naver.log(pid, "{} 클릭 완료".format(searchDate[date]))
        except Exception as e:
            naver.doExcept(pid, store, driver, "{} 클릭 실패".format(searchDate[date]), e)
            return
        
        naver.delay(3)


    # 검색 버튼 클릭
    flag = naver.searchData(pid, driver, task, naver.data[status]["search"])
    if not flag:
        return

    naver.delay(3)
    try:
        # 데이터 검색 후 조회 내역의 두번째 컬럼이 인식되는지로 데이터 존재여부 확인
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, naver.data[status]["data"])))
    except Exception:
        naver.log(pid, "{} 조회 결과 데이터가 존재하지 않음.".format(task))
        return

    # 작업 폴더 비우기(파일 이름 변경 후 결과폴더로 이동해서 원래 폴더에는 파일이 없어야하는데 혹시라도 안지워진게 있으면 문제가 되서 한번더 폴더 안의 파일들 삭제)
    naver.initDir(pid)
        
    flag = naver.downloadExcel(pid, driver, task, naver.data[status]["download"])
    if not flag:
        return

    limit = 60
    path = "{}{}".format(downPath, pid)
    isDownOk = naver.downloadWait(pid, driver, task, store, limit, path, ["Settle"])

    # 몇초 걸려서 기다렸는데도 경로에 파일이 다운되지 못한 경우 한번 더 체크
    isDownOk = naver.downloadCheck(path, ["Settle"])
    if isDownOk == -1:
        naver.storeExcept(pid, driver, task, "{} 파일이 다운로드 되지 못함".format(store))
        return

    # 여기 오면 파일이 정상 다운된 상황    ["정산예정일", "정산완료일", "정산기준일"]
    fnameDict = {"Settle":""}
    fnameDict["Settle"] = naver.data["statusList"][status]
    naver.moveFile(pid, driver, task, "{}{}".format(downPath, pid), downPath, store, fnameDict)

def main(stores):
    # 사이트 접속, 드라이버 생성, 일감 분장, 로그인
    driver = naver.initProcess(downPath, downPath_win, shop, stores)
    # 메인 비즈니스 로직
    doProcess(driver, stores)
    driver.quit()

#멀티프로세싱은 여기서만 조작 가능            #######################################################################
if __name__ == '__main__':
    # 멀티 프로세싱 설정 -> 코어 수만큼 활용, 일감은 csv 파일에서 읽기    ... 코어 수를 4 초과하게 되면 홈페이지가 로봇으로 인식해서 막아버리는 이슈 발생
    num_cores = 4#multiprocessing.cpu_count()
    pool = multiprocessing.Pool(num_cores)

    err = False
    try:
        stores = api.divideWork("{}{}{}".format(naver.nowPath, api.delimeter, naver.data["workFileName"][shop]), num_cores)
    except Exception as e:
        err = True
        naver.adminLog.error("{}: 일감 분리 실패(일감 파일 읽기 불가능)".format(e))

    pool.map(main, stores)
    pool.close()

    line = 0
    day = "{}{}{}".format(api.getYear(), api.getMonth(), api.getDay())
    filename = r"{}_{}_Report.csv".format(day, shop)
    filenames = ["{}".format(filename)]

    filename = r"{}{}log{}{}".format(downPath_win, api.delimeter, api.delimeter, filename)
    home = "/home/rpa01/rpa_naver_brandmall"
    # Windows면 경로 변경
    if api.delimeter == "\\":
        home = "."
    try:
        os.system('{}/sftp.sh 2> {}/sftp.log {}'.format(home, home, filename))
        naver.adminLog.info("{}파일 정상적으로 sftp전송 완료".format(filename))
    except Exception as e:
        naver.adminLog.error("{}파일 정상적으로 sftp전송 실패 | {}".format(filename, e))
    naver.delay(5)

    msg = naver.data["emailText"][task]
    for dir in dirs:
        line = 0
        
        # 웹메일에서 첨부 메일명에 한글이 포함되면 첨부가 되지 않아서 영어 이름으로 매칭
        filename = r"{}_{}_{}.xlsx".format(day, shop, fnames[dir])
        filenames.append('{}'.format(filename))
        filename = r"{}{}{}".format(downPath_win, api.delimeter, filename)

        try:
            excel_concat.getResultFile(r"{}{}{}".format(downPath_win, api.delimeter, dir), filename, line, naver.adminLog, naver.userLog)
            naver.adminLog.info("{}파일 정상적으로 생성 완료".format(dir))
        except Exception:
            err = True
            naver.adminLog.error("{}파일 정상적으로 생성 실패".format(dir))

        isFileExist = False
        try:
            isFileExist = os.path.isfile(filename)
        except Exception as e:
            err = True
            naver.adminLog.error("네이버 정산 {}파일이 존재하지 않음 | {}".format(filename, e))

        # 파일이 정상적으로 생성되어 존재하면, 결과파일 메일에 첨부 파일명 등록
        if isFileExist:
            try:
                os.system('{}/sftp.sh 2> {}/sftp.log {}'.format(home, home, filename))
                naver.adminLog.info("{}파일 정상적으로 sftp전송 완료".format(filename))
            except Exception as e:
                naver.adminLog.error("{}파일 정상적으로 sftp전송 실패 | {}".format(filename, e))
            
            naver.delay(5)
        else:
        # 파일이 존재하지 않으면 결과파일 메일 내용에 기재
            msg = "<br>◈{}{}파일이 정상적으로 생성되지 못하였습니다!!! <br>".format(msg, fnames[dir])

    params = {
        'subject':"({}) 네이버 정산_{}".format(day, shop),
        'to':naver.data["email"][shop],
        'to':'',
        'msg':msg,
        'files':filenames
    }
    # 메일 수신처 설정
    # to = ["chosm10@hyundai-ite.com", "cindy@hyundaihmall.com", "move@hyundai-ite.com"]
    # to.append(naver.data["email"][shop])

    # last_month = api.getLastDate()[0].split(".")[1]
    # try:
    #     mail.sendmail(to, "({}) {}월 네이버 정산_{}".format(day, last_month, shop), msg, files)
    #     naver.adminLog.info("네이버 정산 메일 정상 발송 완료")
    # except Exception as e:
    #     err = True
    #     naver.adminLog.error("네이버 정산 메일 정상 발송 실패 | {}".format(e))

    naver.delay(10)
    api.send_drm(naver.drm_server_ip, params)

    if err:
        data = {'name': task_name, 'botId': bot_id, 'botIp': bot_ip, 'status':'fail'}
    else:
        data = {'name': task_name, 'botId': bot_id, 'botIp': bot_ip, 'status':'comp'}

    try:
        api.post_api(naver.task_status_url, data)
    except Exception as e:
        naver.log(0, e)

    api.taskkill()