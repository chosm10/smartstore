import multiprocessing
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import csv
import json
import api
from os.path import getsize
from os import remove
import shutil
import sys
import os
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

# cmd 실행 할 때 받는 매개변수 -> 점포명   ddp = 동대문, gimpo = 김포
member = sys.argv[1]

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
    comm += member + "/"
    api.mkdir(comm)
    global downPath
    downPath = comm

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
    workFileName = data["workFileName"][member]

def main(stores):
    pid = stores[0]
    log(pid, "프로세스 생성")
    # brands[0] 에는 일감을 받은 프로세스에 할당될 번호가, [1] 부터는 [브랜드명]의 형태로 들어가 있음
    stores = stores[1:]
    logger.info("----------------------------------------------------------------------------------")
    log(pid, "담당 브랜드 목록")
    logger.info(stores)
    logger.info("----------------------------------------------------------------------------------")
    #PID 번호별로 다운로드 받을 작업 디렉터리 생성
    api.mkdir("{}{}".format(downPath, pid))
    #크롬 드라이버 설정
    options = setDownloadPath(pid, 1)
    # 크롬 드라이버 생성
    driver = getDriver(pid, data["driver"], options)
    # 네이버 스마트 스토어 접속
    openStore(pid, driver, data["storeLink"])
    ########################   로그인창까지 링크로 바로 들어갈 수 있어서 주석 처리, 추후 개발 건에서 판매자 아이디 로그인하는 곳은 링크가 달라 분기 처리 필요
    # 로그인 하기 클릭
    # clickMainLogin(pid, driver, data["mainLogin"])
    # 네이버 아이디 로그인 클릭
    # clickIdLogin(pid, driver, data["idLogin"])
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
    doProcess(pid, driver, stores)
    
    #최종적으로 드라이버 닫기
    # driver.close()

#멀티프로세싱은 여기서만 조작 가능
if __name__ == '__main__':
    # 멀티 프로세싱 설정 -> 코어 수만큼 활용, 일감은 csv 파일에서 읽기    ... 코어 수를 4 초과하게 되면 홈페이지가 로봇으로 인식해서 막아버리는 이슈 발생
    num_cores = 4#multiprocessing.cpu_count()
    pool = multiprocessing.Pool(num_cores)
    brands = api.divideWork("{}\\{}".format(mypath, workFileName), num_cores)

    pool.map(main, brands)
    pool.close()
    # 실행 중 사용한 프로세스 정리
    # os.system('taskkill /f /im chrome.exe')
    # os.system('taskkill /f /im python.exe')

# 홈페이지에서 다운로드 버튼 클릭시 저장 될 경로 설정
def setDownloadPath(pid, mode):
    options = webdriver.ChromeOptions()
    # 모드가 0이면 크롬 headless(백엔드 실행)모드 설정
    if mode == 0:
        options.add_argument('headless')
        options.add_argument('--disable-gpu')

    options.add_experimental_option("prefs", {"download.default_directory": r"{}\{}{}\{}{}{}\{}\{}".format(r"C:\Users\Administrator\Desktop\naver\결과파일", api.getYear(), api.getMonth(), api.getYear(), api.getMonth(), api.getDay(), member, pid)
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
            doExcept(pid, "", driver, "클릭 실패: ", e)
    elif func == "key":
        try:
            obj.send_keys(args["key"])
            log(pid, "키 입력 성공")
        except Exception as e:
            doExcept(pid, "", driver, "키 입력 실패: ", e)

# 스토어 이동 클릭하는 함수
def clickMoveStore(pid, store, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 10)
        log(pid, "스토어 이동 버튼 클릭 완료")
    except Exception as e:
        doExcept(pid, store[0], driver, "스토어 이동 버튼 클릭 실패: ", e)

class cannotMoveStoreException(Exception):
    def __init__(self):
        super().__init__('해당 홈페이지로 이동 불가')

class cannotFindStoreException(Exception):
    def __init__(self):
        super().__init__('스토어 리스트에서 스토어명 찾기 불가능')

class cannotReadStoreListException(Exception):
    def __init__(self):
        super().__init__('스토어 리스트 클래스로 인식 불가')

# 원하는 스토어 클릭하는 함수
def clickStore(driver, target):
    # 스토어 이동 눌렀을 때 나오는 팝업창의 스토어들을 클래스로 가져와서 담기
    try:
        stores = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "text-title")))
    except Exception:
        raise cannotReadStoreListException
    for store in stores:
        if target in store.text:
            try:
                driver.execute_script("arguments[0].click();", store) 
            except Exception:
                raise cannotMoveStoreException
            return
    
    # 위에 반복문 전부 돌고 탈출하면 우리가 찾는 스토어가 홈페이지 스토어리스트 상에 없다는 뜻이므로 예외처리
    raise cannotFindStoreException

# 각 스토어별 일처리 하는 함수
def doProcess(pid, driver, stores):
    # 모든 일감들에 대해서
    for store in stores:
        try:
            driver.get("https://sell.smartstore.naver.com/#/home/dashboard")
            delay(3)
        except Exception as e:
            doExcept(pid, store[0], driver, "{} 기본화면으로 이동 불가".format(store[0]), e)
            continue

        # 스토어 이동 클릭
        try:
            clickMoveStore(pid, store, driver, data["moveStore"])
        except Exception as e:
            doExcept(pid, store[0], driver, "{} 스토어 이동 클릭 실패".format(store[0]), e)
            continue

        delay(3)
        # 스토어 클릭
        try:
            clickStore(driver, store[0])
        except cannotMoveStoreException as e:
            doExcept(pid, store[0], driver, "{} 스토어 클릭 실패".format(store[0]), e)
            continue
        except cannotFindStoreException as e:
            doExcept(pid, store[0], driver, "{} 스토어를 홈페이지 리스트에서 찾을 수 없습니다.".format(store[0]), e)
            continue
        except cannotReadStoreListException as e:
            doExcept(pid, store[0], driver, "{} 스토어 이동 창에서 스토어 리스트를 클래스로 인식할 수 없습니다.".format(store[0]), e)
            continue

        # 사이트 이동을 하게 되면 사이트 이동은 안되고 팝업창이 1개 사라지는것으로 확인됨, 그래서 사이트 이동 전에 팝업창 없애기 위해 사이트이동을
        # 하나 더 추가함 / 추후에 팝업창 여러개 발생하면 없애기 위해서 홈페이지 이동을 여러개 추가 할 수 있음 
        try:
            driver.get(data["url"]["cancle"])
            delay(2)
            driver.get(data["url"]["cancle"])
            delay(3)
        except Exception as e:
            cancleAlert(pid, driver)
        # status -> 반품, 취소, 발주발송
        for status in data["url"]:
            # 사이트 이동이 된 후에 첫번째 취소관리로 넘어가야 정상 이동 가능해서 '스마트스토어센터가' 떳는지 확인
            # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["isLogin"])))
            try:
                log(pid, "{} {} 업무 시작".format(store[0], status))
                cancleAlert(pid, driver)
                try:
                    # 페이지 이동 중간에 고객센터 팝업창 떠서 이동 안되는 경우가 있어 팝업창 없애기 용으로 페이지 이동 한번 더함
                    driver.get(data["url"][status])
                    delay(2)
                    driver.get(data["url"][status])
                    delay(2)
                    driver.get(data["url"][status]) #################################혹시 모를 팝업에 대비해 한개 더 추가!!!!!!!!!!!!!!!!!!!!!!!!!
                except Exception as e:
                    storeExcept(pid, "{} {}".format(store[0], data["statusList"][status]), "업무 홈페이지로 이동 불가")
                    continue

                delay(3)
                try:
                    detailJob(pid, driver, store[0], status)
                except Exception as e:
                    logger.error("detailJob 에러 발생")
                    cancleAlert(pid, driver)
                    continue
            except Exception as e:
                logger.error("{} {}에서 처리 안된 에러 발생".format(store[0], status))
                cancleAlert(pid, driver)
                continue

#반품,취소관리, 발주발송 별 세부 엑셀 다운 기능
def detailJob(pid, driver, store, status):
    task = "{} {}".format(store, data["statusList"][status])
    try:
        driver.switch_to_frame(0)
        log(pid,"프레임 변경")
    except Exception:
        log(pid,"프레임 변경 불필요")

    #반품, 취소관리에만 완료 설정 진행
    delay(3)
    if status == "cancle" or status == "return":
        # 3개월 버튼 클릭 (반품, 취소는 3개월)
        try:
            processByXpath(pid, driver, {"xpath":data[status]["searchRange"]}, "click", 5)
            log(pid, "3개월 버튼 클릭 성공")
        except Exception as e:
            storeExcept(pid, task, "3개월 버튼 클릭 실패")
            # 인식 안될 경우 앞에서 팝업창 때문에 홈페이지 이동이 안된 것이므로 작업 종료하고 다음 작업 진행
            return

        delay(3)
        # 처리상태(반품완료, 취소완료) 선택
        try:
            select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,data[status]["processStatus"]))))
            delay(2)
            select.select_by_visible_text(data["statusList"][status])
            log(pid, "처리상태 버튼 클릭 성공")
        except Exception as e:
            storeExcept(pid, task, "처리상태 버튼 클릭 실패")
            return
    elif status == "delivery":
        try:
            select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,data[status]["searchRange"]))))
            delay(2)
            select.select_by_visible_text("발송처리일")
            log(pid, "발송처리일 선택 성공")
        except Exception as e:
            storeExcept(pid, task, "발송처리일 선택 실패")
            return
        
    delay(3)
    # 검색 버튼 클릭
    try:
        processByXpath(pid, driver, {"xpath":data[status]["search"]}, "click", 5)
        log(pid, "검색 버튼 클릭 성공")
    except Exception as e:
        storeExcept(pid, task, "검색 버튼 클릭 실패")
        return

    cnt = 0
    check = False
    while cnt < 6:
        delay(1)
        # 조회 목록이 0이면 바로 다음작업 진행
        try:
            # 정상적인 상황이면 무조건 인식이 되므로 명시적 지연시간 줘도 그만큼 소요 안됨, 인식 안되면 홈페이지 이상
            obj = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, data["dataCnt"])))
        except Exception as e:
            storeExcept(pid, task, "조회 목록 인식 실패")
            return
        # 6초 전에 조회 결과 뜨면 딜레이 멈추고 진행
        if obj.text != "0":
            check = True
            log(pid, "조회목록 {}건 발생".format(obj.text))
            break
        cnt += 1

    # 다운할 양이 0건이면 다음작업 안하고 진행
    if check is False:
        log(pid, "{} 다운 할 데이터 없음(0건)".format(task))
        return

    # 리스트 0번에 뜰때까지 기다림 -> 타임아웃되면 너무 많아서 늦게뜬 경우임 -> 타임아웃이면 넘어감
    try:
        processByXpath(pid, driver, {"xpath":data[status]["checkList"]}, "click", 10)
        log(pid, "조회 항목 로딩 완료")
    except Exception as e:
        storeExcept(pid, task, "조회 항목 로딩이 지연돼 다운 불가(10초 초과)")
        return

    # 조회목록이 존재하면 폴더 비우고 엑셀 다운
    delay(2)
    try:
        # 작업 폴더 비우기(파일 이름 변경 후 결과폴더로 이동해서 원래 폴더에는 파일이 없어야하는데 혹시라도 안지워진게 있으면 문제가 되서 한번더 폴더 안의 파일들 삭제)
        for fname in os.listdir("{}{}".format(downPath, pid)):
            fname = "{}{}/{}".format(downPath, pid, fname)
            try:
                remove(fname)
                logger.info("{} 파일 삭제 완료".format(fname))
            except Exception as e:
                logger.error("{} 파일 삭제 불가".format(fname))


        processByXpath(pid, driver, {"xpath":data[status]["download"]}, "click", 5)
        log(pid, "엑셀다운 버튼 클릭 성공")
        try:
            waitSec = downloadWait("{}{}".format(downPath, pid))
            log(pid, "다운로드 기간 {}초 소요!".format(waitSec))
        except Exception as e:
            storeExcept(pid, task, "다운로드 제한시간 초과!")
            return
        
        # 몇초 걸려서 기다렸는데도 경로에 파일이 다운되지 못한 경우
        check = False
        for fname in os.listdir("{}{}".format(downPath, pid)):
            if "스마트" in fname:
                check = True

        if check is False:
            storeExcept(pid, task, "{} 파일이 다운로드 되지 못함".format(store))
            return

        # 여기 오면 파일이 정상 다운된 상황    
        try:
            moveFile("{}{}".format(downPath, pid), store)
            log(pid, "{} 파일명 변경 및 이동 완료".format(store))
        except Exception as e:
            logger.error(e)
            storeExcept(pid, task, "파일명 및 저장 경로가 정확하지 않음 / 해당 경로에 이미 동일한 파일 존재!!!")
            return
    except Exception as e:
        storeExcept(pid, task, "엑셀다운 버튼 클릭 실패")    
        return
    
def cancleAlert(pid, driver):
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
        except Exception:
            failLog.info("{} 파일 수기 다운 필요".format(name))
            logger.error("{}/{} 파일이 존재하지 않아 이름 변경 불가".format(src, fname))
            return

        # 파일 크기가 0이면 로그 기록 후 삭제하고 리턴
        if getsize("{}/{}".format(src, name)) == 0:
            try:
                remove("{}/{}".format(src, name))
                failLog.info("{} 파일 크기가 0 입니다. 수기 확인 필요".format(name))
                logger.info("{} 파일 크기가 0이여서 삭제".format(name))
            except Exception:
                logger.error("{}/{} 파일을 삭제하지 못했습니다!".format(src, name))
            return

        # 결과파일 폴더로 파일 이동하기
        try:
            shutil.move("{}/{}".format(src, name), target)
            logger.info("{}/{} 파일을 {}으로 이동 완료".format(src, name, target))
        except Exception:
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
        doExcept(pid, "", driver, "{}번 크롬 드라이버 설정 실패: ".format(pid), e)

    return driver

# 네이버 스마트 스토어 접속
def openStore(pid, driver, link):
    try:
        driver.get(link)
        log(pid, "스마트 스토어 홈페이지 접속 완료")
    except Exception as e:
        doExcept(pid, "", driver, "스마트 스토어 홈페이지 접속 실패: ", e)

    try:    
        driver.maximize_window()
        log(pid, "홈페이지 최대화 성공")
    except Exception as e:
        doExcept(pid, "", driver, "홈페이지 최대화 실패: ", e)

# 여기에 걸리면 한 프로세서 자체가 일을 못한것이므로 해당 일감 전체 수기 다운 필요
def doExcept(pid, store, driver, msg, e):
    cancleAlert(pid, driver)
    failLog.info("{} 3건(반품, 취소, 발주발송)".format(store))
    logger.error("PID: {} | {}".format(pid, msg))
    logger.error(e)

# 여기에 걸리면 해당 브랜드의 한 작업(반품 / 취소/ 발주발송)에 대해 수기 다운 필요
def storeExcept(pid, task, msg):
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
        doExcept(pid, "", driver, "스마트 스토어 첫 접속페이지 로그인하기 버튼 인식 실패: ", e)

# 네이버 아이디 로그인 클릭
def clickIdLogin(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 30)
        log(pid, "네이버 아이디로 로그인하기 버튼 인식 성공")
    except Exception as e:
        doExcept(pid, "", driver, "네이버 아이디로 로그인하기 버튼 인식 실패: ", e)

# 아이디 입력
def inputId(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath": xpath, "key": data["members"][member][0]}, "key", 30)
        log(pid, "아이디 입력 칸 인식 성공")
    except Exception as e:
        doExcept(pid, "", driver, "아이디 입력 칸 인식 실패: ", e)

# 비밀번호 입력
def inputPwd(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath": xpath, "key": data["members"][member][1]}, "key", 30)
        log(pid, "비밀번호 입력 칸 인식 성공")
    except Exception as e:
        doExcept(pid, "", driver, "비밀번호 입력 칸 인식 실패: ", e)
    
# 마지막 로그인버튼 클릭
def clickLastLogin(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 10)
        log(pid, "로그인 버튼 인식 성공")
    except Exception as e:
        doExcept(pid, "", driver, "로그인 버튼 인식 실패: ", e)

# 정상 로그인 여부 확인
def isLoginOk(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "",10)
        log(pid, "정상 로그인 완료")
    except Exception as e:
        doExcept(pid, "", driver, "로그인 불가: ", e)

# 등록안함 버튼 클릭
def clickNotReg(pid, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 30)
        log(pid, "등록안함 버튼 인식 성공")
    except Exception as e:
        logger.error("등록안함 버튼 인식 실패, xpath 및 계정정보 변경 확인 필요!!!: {}".format(e))

def delay(sec):
    time.sleep(sec)