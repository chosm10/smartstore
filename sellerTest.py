import multiprocessing
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import csv
import json
from comm import api
from os.path import getsize
from os import remove
import shutil
import sys
import os
from selenium.webdriver.common.keys import Keys
# 실행 시 cmd 창 맨위에 놓으면 크롤링 시 버튼 인식 안되니 cmd 창 내려놓고 실행해야함
###########################################################################################################################################################################
# 1. RPA에서는 안되는 멀티 프로세싱 기능을 적용해서 이론 상 RPA 보다 해당 PC의 코어수 만큼 병렬처리(동시에 실행)하여 시간 단축 가능                                         
# 2. 처음에는 팝업창을 제거 해줘야만 그 뒤에 있는 화면의 버튼 요소들이 인식 가능한 줄 알았음 -> 여러개 팝업창이 뜰 가능성이 있어서 모든 팝업창에 대해 대기시간을 놓고 인식을
#    대기하는 과정이 필요 했음 -> 개발 중 팝업창을 제거하지 않아도 frame 순서만 변경해주면 팝업창 떠있는 상태로 조작 가능함을 알게됨 -> 시간 단축
# 3. 현재 RPA로 개발된 프로그램보다 코드 길이가 훨씬 짧음 -> 기존에는 홈페이지 변경 시 프로그램마다 변경 필요했으나 추상화를 적용해서 유지보수 시에도 공통 영역의 코드 변경 시 일괄 적용이 가능해짐
# 4. 위 내용을 종합적으로 시간 단축 가능, RPA 라이선스 미사용으로 금액 절감 가능
# 5. 파이썬 적용 시 파이썬을 배워야하고, 관리 측면에서 RPA와 파이썬을 둘 다 관리해야 한다는 단점이 있긴 함
###########################################################################################################################################################################
commDir = "C:/Users/Administrator/Desktop/naver/"
downPath = "{}결과파일/{}{}/{}{}{}/".format(commDir, api.getYear(), api.getMonth(), api.getYear(), api.getMonth(), api.getDay())
errors = []
mypath = os.path.dirname(os.path.realpath(__file__))
logger = None
failLog = None
def mkdir():
    dirs = ["반품완료", "취소완료", "발주발송(발송처리일)"]
    comm = commDir
    comm += "결과파일/"
    api.mkdir(comm)
    day = api.getYear() + api.getMonth()
    comm += day + "/"
    api.mkdir(comm)
    day += api.getDay()
    comm += day + "/"
    api.mkdir(comm)

    for d in dirs:
        api.mkdir(comm + d)

    #로그 폴더 생성
    api.mkdir("{}log".format(comm))
    # 로거 설정
    logFileName = "{}/log/{}{}{}.log".format(comm, api.getYear(), api.getMonth(), api.getDay())
    failFileName = "{}/log/{}{}{}_보고파일.csv".format(comm, api.getYear(), api.getMonth(), api.getDay())

    global logger
    logger = api.getLogger(logFileName, "naver_daily_selling")
    global failLog
    failLog = api.failHistory(failFileName, "naver_daily_fail")
    failLog.info("***********재다운로드 필요 항목***********")

mkdir()
with open(mypath +'\\data.json', encoding='utf-8') as json_file:
    data = json.load(json_file)        
    # img capute path 설정
    imgPath = data["imgPath"] 
    workFileName = data["workFileName"]

# cmd 실행 할 때 받는 매개변수 -> 점포명   ddp = 동대문, gimpo = 김포
member = sys.argv[1]

def main(stores):
    pid = stores[0]
    log(pid, "프로세스 생성")
    # brands[0] 에는 일감을 받은 프로세스에 할당될 번호가, [1] 부터는 [브랜드명]의 형태로 들어가 있음
    stores = stores[1:]
    logger.info("----------------------------------------------------------------------------------")
    log(pid, "담당 브랜드 목록")
    logger.info(stores)
    logger.info("----------------------------------------------------------------------------------")
    api.mkdir("{}{}".format(downPath, pid))
    options = setDownloadPath(pid)
    # 크롬 드라이버 생성
    driver = getDriver(pid, data["driver"], options)
    # 네이버 스마트 스토어 접속
    openStore(pid, driver, data["storeLink"])
    # 로그인 하기 클릭
    clickMainLogin(pid, driver, data["mainLogin"])
    # 네이버 아이디 로그인 클릭
    clickIdLogin(pid, driver, data["idLogin"])
    # 아이디 입력
    inputId(pid, driver, data["idBox"])
    delay(3) # 타이핑 지연시간
    # 비밀번호 입력
    inputPwd(pid, driver, data["pwdBox"])
    delay(3)
    # 로그인 버튼 클릭
    clickLastLogin(pid, driver, data["lastLogin"])
    # 동대문점은 로그인 후 등록안함 버튼을 눌러줘야함
    global member
    if member == "ddp":
        clickNotReg(pid, driver, data["notReg"])
    # 정상 로그인 되어 스마트스토어센터 글자가 인식 되는지 확인
    isLoginOk(pid, driver, data["isLogin"])    
    # 메인 비즈니스 로직
    doProcess(pid, driver)
    
    #최종적으로 드라이버 닫기
    # driver.close()

#멀티프로세싱은 여기서만 조작 가능
if __name__ == '__main__':
    # 멀티 프로세싱 설정 -> 코어 수만큼 활용, 일감은 csv 파일에서 읽기    ... 코어 수를 4 초과하게 되면 홈페이지가 로봇으로 인식해서 막아버리는 이슈 발생
    num_cores = 1#multiprocessing.cpu_count()
    pool = multiprocessing.Pool(num_cores)
    brands = api.divideWork("{}\\{}".format(mypath, workFileName), num_cores)

    pool.map(main, brands)
    pool.close()
    # os.system('taskkill /f /im chrome.exe')

# 홈페이지에서 다운로드 버튼 클릭시 저장 될 경로 설정
def setDownloadPath(pid):
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {"download.default_directory": r"{}\{}{}\{}{}{}\{}".format(r"C:\Users\Administrator\Desktop\naver\결과파일", api.getYear(), api.getMonth(), api.getYear(), api.getMonth(), api.getDay(), pid)
    })
    return options

# xpath로 클릭, 키보드 입력하는 함수, args에 필요 정보를 받음
def processByXpath(pid, driver, args, func, delay):
    # 명시적으로 해당 요소가 나타날때 까지 delay초 기다림
    obj = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH,args["xpath"])))

    if func == "click":
        try:
            driver.execute_script("arguments[0].click();", obj)
            log(pid, "클릭 성공")
        except Exception as e:
            doExcept(pid, driver, "클릭 실패: ", e)
    elif func == "key":
        try:
            obj.send_keys(args["key"])
            log(pid, "키 입력 성공")
        except Exception as e:
            doExcept(pid, driver, "키 입력 실패: ", e)

# 스토어 이동 클릭하는 함수
def clickMoveStore(pid, store, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 10)
        log(pid, "스토어 이동 버튼 클릭 완료")
    except Exception as e:
        doExcept(pid, store[0], driver, "스토어 이동 버튼 클릭 실패: ", e)

# 원하는 스토어 클릭하는 함수
def clickStore(driver, target):
    # 스토어 이동 눌렀을 때 나오는 팝업창의 스토어들을 클래스로 가져와서 담기
    stores = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "text-title")))
    for store in stores:
        if target in store.text:
            driver.execute_script("arguments[0].click();", store) 
            return

# 각 스토어별 일처리 하는 함수
def doProcess(pid, driver):
    # 모든 일감들에 대해서
    try:
        driver.get("https://sell.smartstore.naver.com/#/home/dashboard")
        delay(3)
    except Exception as e:
        doExcept(pid, "ddp", driver, "{} 기본화면으로 이동 불가".format("ddp"), e)

        # 사이트 이동을 하게 되면 사이트 이동은 안되고 팝업창이 1개 사라지는것으로 확인됨, 그래서 사이트 이동 전에 팝업창 없애기 위해 사이트이동을
        # 하나 더 추가함 / 추후에 팝업창 여러개 발생하면 없애기 위해서 홈페이지 이동을 여러개 추가 할 수 있음 
    try:
        driver.get(data["url"]["account"])
        delay(2)
        driver.get(data["url"]["account"])
        delay(3)
    except Exception as e:
        cancleAlert(pid)

    try:
        log(pid, "{} {} 업무 시작".format("ddp", "account"))
        logger.info("여긴옴1")
        cancleAlert(pid)
        logger.info("여긴옴2")
        try:
            # 페이지 이동 중간에 고객센터 팝업창 떠서 이동 안되는 경우가 있어 팝업창 없애기 용으로 페이지 이동 한번 더함
            driver.get(data["url"]["account"])
            delay(2)
            driver.get(data["url"]["account"])
            delay(2)
            driver.get(data["url"]["account"])
        except Exception as e:
            storeExcept(pid, "{} {}".format("ddp", data["statusList"]["account"]), "업무 홈페이지로 이동 불가")

        delay(3)
        try:
            detailJob(pid, driver)
        except Exception as e:
            logger.error("detailJob 에러 발생")
            cancleAlert(pid)
    except Exception as e:
        cancleAlert(pid)    

#반품,취소관리, 발주발송 별 세부 엑셀 다운 기능
def detailJob(pid, driver):
    try:
        driver.switch_to_frame(0)
        log(pid,"프레임 변경")
    except Exception:
        log(pid,"프레임 변경 불필요")

    # 첫번째 날짜 칸 인식
    obj = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='startDate']/input")))
    # 읽기 전용인 기능을 없애줌
    driver.execute_script('arguments[0].removeAttribute("readonly")', obj)

    try:
        # 이전에 적혀있던 날짜 삭제
        obj.send_keys(Keys.CONTROL + "a" + Keys.DELETE)
        obj.send_keys("2020.07.06")
        log(pid, "키 입력 성공")
    except Exception as e:
        doExcept(pid, driver, "키 입력 실패: ", e)

    obj = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,"//*[@id='searchForm']/div[3]/button[1]")))
    driver.execute_script("arguments[0].click();", obj)
    
def cancleAlert(pid):
    # 엑셀 다운로드 눌렀는데 경고창 발생 시 닫고 넘어감(경고창 2번발생하는 경우도 있어 2개처리)
    try:
        result = driver.switch_to_alert()
        log(pid, result.text)
        result.accept()
        result.dismiss()
    except Exception:
        log(pid, "경고창 발생 안함")
        return
    try:
        result = driver.switch_to_alert()
        log(pid, result.text)
        result.accept()
        result.dismiss()
    except Exception:
        log(pid, "경고창 발생 안함")

def downloadWait(path):
    wait = True
    second = 0
    while wait and second < 30:
        delay(1)
        for fname in os.listdir(path):
            if fname.startswith("스마트"):
                wait = False
        second += 1

    if second >= 30:
        raise Exception
    
    return second

def moveFile(src, store):
    for fname in os.listdir(src):
        if fname == "":
            continue

        target = ""
        name = ""
        if "반품관리" in fname:
            name = "{}_{}.xlsx".format(store, "반품완료")
            target = "{}{}".format(downPath, "반품완료")
        elif "취소관리" in fname:
            name = "{}_{}.xlsx".format(store, "취소완료")
            target = "{}{}".format(downPath, "취소완료")
        elif "전체주문조회" in fname:
            name = "{}_{}.xlsx".format(store, "발주발송(발송처리일)")
            target = "{}{}".format(downPath, "발주발송(발송처리일)")
        else:
            logger.error("이름이 공백, 반품관리, 취소관리, 전체주문조회 모두 아닌 경우 발생: {}".format(fname))
            continue

        # Default 이름 형식에서 업무에 적용되는 이름 형식으로 변경
        try:
            os.rename("{}/{}".format(src, fname), "{}/{}".format(src, name))
            logger.info("파일명을 {}/{}에서 {}/{}로 변경 완료".format(src, fname, src, name))
        except Exception as e:
            failLog.info("{} 파일 수기 다운 필요".format(name))
            logger.error("{}/{} 파일이 존재하지 않아 이름 변경 불가".format(src, fname))
            return

        # 파일 크기가 0이면 로그 기록 후 삭제하고 리턴
        if getsize("{}/{}".format(src, name)) == 0:
            try:
                remove("{}/{}".format(src, name))
                failLog.info("{} 파일 크기가 0 입니다. 수기 확인 필요".format(name))
                logger.info("{} 파일 크기가 0이여서 삭제".format(name))
            except Exception as e:
                logger.error("{}/{} 파일을 삭제하지 못했습니다!".format(src, name))
            return

        # 결과파일 폴더로 파일 이동하기
        try:
            shutil.move("{}/{}".format(src, name), target)
            logger.info("{}/{} 파일을 {}으로 이동 완료".format(src, name, target))
        except Exception as e:
            failLog.info("{} 파일 수기 다운 필요".format(name))
            logger.error("{}/{} 파일 이름이 폴더에 존재하지 않음".format(src, name))
        # 파일이 옮겨지기 전에 또 읽어버리는 이슈가 있어 1번만 수행하고 리턴
        return

# 크롬드라이버 연결
def getDriver(pid, path, option):
    try:
        driver = webdriver.Chrome(path + 'chromedriver', chrome_options=option)
        log(pid, "크롬 드라이버 설정 완료")
    except Exception as e:
        doExcept(pid, driver, "{}번 크롬 드라이버 설정 실패: ".format(pid), e)

    return driver

# 네이버 스마트 스토어 접속
def openStore(pid, driver, link):
    try:
        driver.get(link)
        log(pid, "스마트 스토어 홈페이지 접속 완료")
    except Exception as e:
        doExcept(pid, driver, "스마트 스토어 홈페이지 접속 실패: ", e)

    try:    
        driver.maximize_window()
        log(pid, "홈페이지 최대화 성공")
    except Exception as e:
        doExcept(pid, driver, "홈페이지 최대화 실패: ", e)

# 여기에 걸리면 한 프로세서 자체가 일을 못한것이므로 해당 일감 전체 수기 다운 필요
def doExcept(pid, store, driver, msg, e):
    cancleAlert(pid)
    failLog.info("{} 3건(반품, 취소, 발주발송)".format(store))
    logger.error("PID: {} | {}".format(pid, msg))
    logger.error(e)

# 여기에 걸리면 해당 브랜드의 한 작업(반품 / 취소/ 발주발송)에 대해 수기 다운 필요
def storeExcept(pid, task, msg):
    cancleAlert(pid)
    failLog.info("{}".format(task))    
    logger.error("PID: {} | {}".format(pid, msg))

def log(pid, msg):
    logger.info("PID: {} | {}".format(pid, msg))

# 스마트스토어 메인 페이지에서 로그인하기 버튼 클릭
def clickMainLogin(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 30)
        log(pid,"스마트 스토어 첫 접속페이지 로그인하기 버튼 인식 성공")
    except Exception as e:
        doExcept(pid, driver, "스마트 스토어 첫 접속페이지 로그인하기 버튼 인식 실패: ", e)

# 네이버 아이디 로그인 클릭
def clickIdLogin(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 30)
        log(pid, "네이버 아이디로 로그인하기 버튼 인식 성공")
    except Exception as e:
        doExcept(pid, driver, "네이버 아이디로 로그인하기 버튼 인식 실패: ", e)

# 아이디 입력
def inputId(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath": xpath, "key": data["members"][member][0]}, "key", 30)
        log(pid, "아이디 입력 칸 인식 성공")
    except Exception as e:
        doExcept(pid, driver, "아이디 입력 칸 인식 실패: ", e)

# 비밀번호 입력
def inputPwd(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath": xpath, "key": data["members"][member][1]}, "key", 30)
        log(pid, "비밀번호 입력 칸 인식 성공")
    except Exception as e:
        doExcept(pid, driver, "비밀번호 입력 칸 인식 실패: ", e)
    
# 마지막 로그인버튼 클릭
def clickLastLogin(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 10)
        log(pid, "로그인 버튼 인식 성공")
    except Exception as e:
        doExcept(pid, driver, "로그인 버튼 인식 실패: ", e)

# 정상 로그인 여부 확인
def isLoginOk(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "",10)
        log(pid, "정상 로그인 완료")
    except Exception as e:
        doExcept(pid, driver, "로그인 불가: ", e)

# 등록안함 버튼 클릭
def clickNotReg(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 30)
        log(pid, "등록안함 버튼 인식 성공")
    except Exception as e:
        logger.error("등록안함 버튼 인식 실패, xpath 및 계정정보 변경 확인 필요!!!: {}".format(e))

def delay(sec):
    time.sleep(sec)