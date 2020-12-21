from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import multiprocessing
import time
import csv
import sys
from comm import naver
from comm import api

# 프로그램 다운 경로
downPath = "C:/Users/Administrator/Desktop/naver/"
# 크롬 다운로드 윈도우 경로
downPath_win = r"C:\Users\Administrator\Desktop\naver"
# cmd 실행 할 때 받는 매개변수 -> 점포명   ddp = 동대문, gimpo = 김포
shop = sys.argv[1]
task = "makebrandlist"
naver.mkdir(downPath, downPath_win, task, shop, [])
def makeBrandList(driver):
    try:
        time.sleep(3)

        objs = driver.find_elements_by_class_name("text-title")
        print("리스트 갯수: ", len(objs))

        f = open('{}/brands_{}.csv'.format(naver.nowPath, shop), 'w', encoding='cp949', newline='')
        wr = csv.writer(f)
        #담아놔도 스토어 선택 창 열면 각 버튼들의 xpath가 달라져서 어차피 인식이 안됨 -> 열때마다 모든 스토어 중에서 찾아서 이동해야함
        for o in objs:
            # target - > [[],[], ...] 
            # 브랜드명 앞의 6글자인 '스마트스토어'를 제외하고 순수 브랜드명만 기록
            if "스마트" in o.text:
                wr.writerow([o.text[6:]])
            elif "백화점" in o.text:
                wr.writerow([o.text[3:]])
        f.close()
        
    except Exception:
        print("스토어리스트 파일 생성 중 에러 발생")

def main(stores):
    # 사이트 접속, 드라이버 생성, 일감 분장, 로그인
    driver = naver.initProcess(naver.downPath, naver.downPath_win, shop, stores)
    try:
        naver.clickMoveStore(0, "", driver, naver.data["moveStore"])
    except Exception:
        print("스토어 이동 클릭 실패")
    makeBrandList(driver)

if __name__ == '__main__':
    num_cores = 1
    pool = multiprocessing.Pool(num_cores)
    try:
        stores = api.divideWork("{}\\{}".format(naver.nowPath, naver.data["workFileName"][shop]), num_cores)
    except Exception as e:
        naver.adminLog.error("{}: 일감 분리 실패(일감 파일 읽기 불가능)".format(e))

    pool.map(main, stores)
    pool.close()