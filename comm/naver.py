from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import csv
import json
from . import api
from . import excel_concat
from os.path import getsize
from os import remove
import shutil
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

# task: 업무명(naverdaily~), shop: 점포명(ddp, gimpo ... ), dirs: 내부 작업(반품완료, 일별정산...)
def mkdir(downP, downP_win, task, shop, dirs):
    # 다운로드 경로 년월일 까지 디렉터리 생성
    global downPath
    downPath = downP
    api.mkdir(downPath)
    downPath += "{}/".format(task)
    api.mkdir(downPath)

    downPath += "{}/".format(shop)
    api.mkdir(downPath)

    downPath += "{}/{}/{}/".format(api.getYear(), api.getMonth(), api.getDay())
    api.mkdir(downPath)

    # 각 업무별 다운로드 저장 디렉터리 생성
    for d in dirs:
        api.mkdir(downPath + d)

    api.mkdir(downPath + "screenshot")
    
    global downPath_win
    downPath_win = downP_win
    downPath_win = r"{}\{}\{}\{}\{}\{}".format(downPath_win, task, shop, api.getYear(), api.getMonth(), api.getDay())

    #로그 폴더 생성
    api.mkdir("{}log".format(downPath))
    # 로거 설정
    adminLogPath = "{}/log/{}{}{}.log".format(downPath, api.getYear(), api.getMonth(), api.getDay())
    userLogPath = "{}/log/{}{}{}_Report.csv".format(downPath, api.getYear(), api.getMonth(), api.getDay())

    global adminLog
    adminLog = api.getAdminLogger(adminLogPath, "{}_selling".format(task))
    global userLog
    userLog = api.getUserLogger(userLogPath, "{}_fail".format(task))
    userLog.info("***********재다운로드 필요 항목***********")

# 현재 py 파일이 있는 경로
nowPath = "{}\\res".format(os.path.dirname(os.path.realpath(__file__))) 
with open(nowPath +'\\data.json', encoding='utf-8') as json_file:
    data = json.load(json_file)     

def initProcess(downPath, downPath_win, shop, stores):
    # pid 번호
    pid = stores[0]
    log(pid, "프로세스 생성")
    # sotres[0] 에는 일감을 받은 프로세스에 할당될 번호가, [1] 부터는 [브랜드명]의 형태로 들어가 있음
    stores = stores[1:]
    adminLog.info("----------------------------------------------------------------------------------")
    log(pid, "담당 브랜드 목록")
    adminLog.info(stores)
    adminLog.info("----------------------------------------------------------------------------------")
    #PID 번호별로 다운로드 받을 작업 디렉터리 생성
    api.mkdir("{}{}".format(downPath, pid))
    #크롬 드라이버 설정
    options = setDriverOption(pid, downPath_win, 0) 
    # 크롬 드라이버 생성
    driver = getDriver(pid, data["driver"], options)

    # 로그인 화면이 "네이버 아이디 로그인"이면 0, "판매자 아이디 로그인" 이면 1
    loginSelect = 0
    if shop == "donggu":
        loginSelect = 1

    # 네이버 스마트 스토어 접속
    openStore(pid, driver, data["storeLink"][loginSelect])
    # 아이디 입력
    inputId(pid, shop, driver, data["idBox"][loginSelect])
    delay(3) # 타이핑 지연시간
    # 비밀번호 입력
    inputPwd(pid, shop, driver, data["pwdBox"][loginSelect])
    delay(3)
    # 로그인 버튼 클릭
    clickLastLogin(pid, driver, data["lastLogin"][loginSelect])
    # 동대문점은 로그인 후 등록안함 버튼을 눌러줘야함
    if shop == "ddp":
        clickNotReg(pid, driver, data["notReg"])
    # 정상 로그인 되어 스마트스토어센터 글자가 인식 되는지 확인
    isLoginOk(pid, driver, data["isLogin"])    

    return driver

# 홈페이지에서 다운로드 버튼 클릭시 저장 될 경로 설정
def setDriverOption(pid, downPath_win, mode):
    options = webdriver.ChromeOptions()
    # 모드가 0이면 크롬 headless(백엔드 실행)모드 설정
    if mode == 0:
        options.add_argument('headless')
        options.add_argument('--disable-gpu')

    options.add_experimental_option("prefs", {"download.default_directory": r"{}\{}".format(downPath_win, pid)
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

def searchData(pid, driver, task, xpath):
    try:
        processByXpath(pid, driver, {"xpath":xpath}, "click", 5)
        log(pid, "검색 버튼 클릭 성공")
    except Exception:
        storeExcept(pid, driver, task, "검색 버튼 클릭 실패")
        return False
    return True

def downloadExcel(pid, driver, task, xpath):
    try:
        processByXpath(pid, driver, {"xpath": xpath}, "click", 5)
        log(pid, "엑셀다운 버튼 클릭 성공")
    except Exception:
        storeExcept(pid, driver, task, "엑셀 다운로드 버튼 클릭 실패")
        return False
    return True

def moveStore(pid, driver, store):
    try:
        driver.get("https://sell.smartstore.naver.com/#/home/dashboard")
        delay(3)
    except Exception as e:
        doExcept(pid, store, driver, "{} 기본화면으로 이동 불가".format(store), e)
        return False

    # 스토어 이동 클릭
    try:
        clickMoveStore(pid, store, driver, data["moveStore"])
    except Exception as e:
        doExcept(pid, store, driver, "{} 스토어 이동 클릭 실패".format(store), e)
        return False

    delay(3)
        # 스토어 클릭
    try:
        clickStore(driver, store)
    except cannotMoveStoreException as e:
        doExcept(pid, store, driver, "{} 스토어 클릭 실패".format(store), e)
        return False
    except cannotFindStoreException as e:
        doExcept(pid, store, driver, "{} 스토어를 홈페이지 리스트에서 찾을 수 없습니다.".format(store), e)
        return False
    except cannotReadStoreListException as e:
        doExcept(pid, store, driver, "{} 스토어 이동 창에서 스토어 리스트를 클래스로 인식할 수 없습니다.".format(store), e)
        return False
    
    return True

# 사이트 이동을 하게 되면 사이트 이동은 안되고 팝업창이 1개 사라지는것으로 확인됨, 그래서 사이트 이동 전에 팝업창 없애기 위해 사이트이동을
# 하나 더 추가함 / 추후에 팝업창 여러개 발생하면 없애기 위해서 홈페이지 이동을 여러개 추가 할 수 있음 
def canclePopup(pid, driver, task, store, status):
    try:
        # 페이지 이동을 1번 할때 마다 팝업창을 1개씩 제거 할 수 있음, 현재 최대 3개의 팝업창이 발생하는것으로 확인됨.
        driver.get(data["url"][task][status])
        delay(2)
        driver.get(data["url"][task][status])
        delay(2)
        driver.get(data["url"][task][status])
        delay(2)
        driver.get(data["url"][task][status])#################################위에 팝업창 제거되고 실체로 페이지 이동되는 부분!!!!!!!!!!!!!!!!!!!!!!!!!
    except Exception:
        storeExcept(pid, driver, "{} {}".format(store, data["statusList"][status]), "업무 홈페이지로 이동 불가")
        cancleAlert(pid, driver)
        return False

    return True

# 프레임 변경이 필요한 경우 변경하는 기능
def switchFrame(pid, driver):
    try:
        driver.switch_to_frame(0)
        log(pid,"프레임 변경")
    except Exception:
        log(pid,"프레임 변경 불필요")

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
    prename = ""
    for store in stores:
        # 스마트스토어 홈페이지 브랜드 리스트에 백화점, 스마트스토어가 혼용되어 추출한 text에 해당 문구가 있는 경우 브랜드명 앞에 붙여서 일치여부 판단
        if store.text.startswith("스마트"):
            prename = "스마트스토어"
        elif store.text.startswith("백화점"):
            prename = "백화점"

        if "{}{}".format(prename, target) == store.text:
            try:
                driver.execute_script("arguments[0].click();", store) 
            except Exception:
                raise cannotMoveStoreException
            return
    
    # 위에 반복문 전부 돌고 탈출하면 우리가 찾는 스토어가 홈페이지 스토어리스트 상에 없다는 뜻이므로 예외처리
    raise cannotFindStoreException

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
# keywords: 업무마다 다운 받는 파일의 default 이름 형식이 달라서 식별하기 위한 문구 ex) 일매출정리 -> "스마트"
def downloadWait(pid, driver, task, store, limit, path, keywords):
    second = 0
    while second < limit:
        delay(1)
        try:
            flag = downloadCheck(path, keywords)
            if flag != -1:
                log(pid, "다운로드 기간 {}초 소요!".format(second))
                return flag
        except Exception as e:
            storeExcept(pid, driver, task, "다운로드 체크 중 오류 발생 | {}".format(e))
            return -1

        second += 1
    # 앞에서 리턴 안되면 다운로드 제한 초과한것임
    storeExcept(pid, driver, task, "다운로드 제한시간 초과")
    return -1

def downloadCheck(path, keywords):
    for fname in os.listdir(path):
        for keyword in keywords:
            if keyword in fname:
                fname = r"{}\{}".format(path, fname)
                isDownOk = excel_concat.fileRowCheck(fname)
                return isDownOk
    return -1

# fNameDict -> {"반품관리": "반품완료", "취소관리": "취소완료", "전체주문": "발주발송(발송처리일)" ...}
def moveFile(pid, driver, task, src, downPath, store, fNameDict):
    try:
        for fname in os.listdir(src):
            if fname == "":
                continue

            target = ""
            name = ""
            flag = False
            for key in fNameDict.keys():
                if key in fname:
                    name = "{}_{}.xlsx".format(store, fNameDict[key])
                    target = "{}{}".format(downPath, fNameDict[key])
                    flag = True
                    break
            
            if not flag:
                storeExcept(pid, driver, task, "이름이 공백이거나 fNameDict대로 다운되지 않은 경우 발생: {}".format(fname))
                continue    

            # Default 이름 형식에서 업무에 적용되는 이름 형식으로 변경
            try:
                os.rename("{}/{}".format(src, fname), "{}/{}".format(src, name))
                log(pid, "파일명을 {}/{}에서 {}/{}로 변경 완료".format(src, fname, src, name))
            except Exception as e:
                storeExcept(pid, driver, task, "{}/{} 파일이 존재하지 않아 이름 변경 불가 | {}".format(src, fname, e))
                return

            # 파일 크기가 0이면 로그 기록 후 삭제하고 리턴
            if getsize("{}/{}".format(src, name)) == 0:
                try:
                    remove("{}/{}".format(src, name))
                    userLog.info("{} 파일 크기가 0 입니다. 수기 확인 필요".format(name))
                    log(pid, "{} 파일 크기가 0이여서 삭제".format(name))
                except Exception as e:
                    storeExcept(pid, driver, task, "{}/{} 파일을 삭제하지 못했습니다! | {}".format(src, name, e))
                return

            # 결과파일 폴더로 파일 이동하기
            try:
                shutil.move("{}/{}".format(src, name), target)
                log(pid, "{}/{} 파일을 {}으로 이동 완료".format(src, name, target))
            except Exception as e:
                storeExcept(pid, driver, task, "{}/{} 파일 이름이 폴더에 존재하지 않음 | {}".format(src, name, e))
            # 파일이 옮겨지기 전에 또 읽어버리는 이슈가 있어 1번만 수행하고 리턴
            log(pid, "{} 파일명 변경 및 이동 완료".format(store))
            return
    except Exception as e:
        storeExcept(pid, driver, task, "파일명 및 저장 경로가 정확하지 않음 / 해당 경로에 이미 동일한 파일 존재!!! | {}".format(e))

# 작업 폴더 비우기(파일 이름 변경 후 결과폴더로 이동해서 원래 폴더에는 파일이 없어야하는데 혹시라도 안지워진게 있으면 문제가 되서 한번더 폴더 안의 파일들 삭제)
def initDir(pid):
    try:
        for fname in os.listdir("{}{}".format(downPath, pid)):
            fname = "{}{}/{}".format(downPath, pid, fname)
            try:
                remove(fname)
                adminLog.info("{} 파일 삭제 완료".format(fname))
            except Exception:
                adminLog.error("{} 파일 삭제 불가".format(fname))
    except Exception as e:
        adminLog.error("{} 디렉터리 경로가 존재하지 않습니다. | {}".format("{}{}".format(downPath, pid), e))

# 크롬드라이버 연결
def getDriver(pid, path, option):
    try:
        driver = webdriver.Chrome(path + 'chromedriver', chrome_options=option)
        log(pid, "크롬 드라이버 설정 완료")
    except Exception as e:
        adminLog.error("{}번 크롬 드라이버 설정 실패: | {}".format(pid, e))
        # 설치된 크롬드라이버 버젼이 안맞게 될 경우 매니저로 다운받아와서 실행
        driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=option)
        return driver

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
    userLog.info("{} 관련 모든 업무".format(store))
    adminLog.error("PID: {} | {} | {}".format(pid, msg, e))
    api.capture(pid, driver, r"{}\{}".format(downPath_win,"screenshot"))

# 여기에 걸리면 해당 브랜드의 한 작업(반품 / 취소/ 발주발송)에 대해 수기 다운 필요
def storeExcept(pid, driver, task, msg):
    userLog.info("{}".format(task))    
    adminLog.error("PID: {} | {}".format(pid, msg))
    api.capture(pid, driver, r"{}\{}".format(downPath_win, "screenshot"))

def log(pid, msg):
    adminLog.info("PID: {} | {}".format(pid, msg))

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
def inputId(pid, shop, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath": xpath, "key": data["members"][shop][0]}, "key", 30)
        log(pid, "아이디 입력 칸 인식 성공")
    except Exception as e:
        doExcept(pid, "", driver, "아이디 입력 칸 인식 실패: ", e)

# 비밀번호 입력
def inputPwd(pid, shop, driver, xpath):
    try:
        processByXpath(pid, driver, {"xpath": xpath, "key": data["members"][shop][1]}, "key", 30)
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
        adminLog.error("등록안함 버튼 인식 실패, xpath 및 계정정보 변경 확인 필요!!!: {}".format(e))

def delay(sec):
    try:
        time.sleep(sec)
    except Exception as e:
        adminLog.error("시간 지연 불가 오류 발생! | {}".format(e))

def setDRM(src):
    try:
        os.system(r'cscript .\comm\setDRM.vbs {}'.format(src)) 
        adminLog.info('{}에 DRM 정상 설정 완료'.format(src))
    except Exception as e:
        adminLog.error('{}에 DRM이 설정되지 못하였습니다. | {}'.format(src, e))