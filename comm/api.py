import os
import time
import logging
import csv
import sys
import errno
from datetime import date, timedelta

def capture(pid, driver, path):
    now = time.localtime()
    timeForm = "pid%s-%04d-%02d-%02d-%02dh-%02dm-%02ds" % (pid, now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
    saveas = r"{}\{}{}".format(path, timeForm, '.png')
    try:
        driver.save_screenshot(saveas)
        print(saveas + "저장 완료")
    except Exception:
        print("스크린샷 에러")

def getAdminLogger(logFileName, name):
    Logger = logging.getLogger(name)
    Logger.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(filename)s - %(lineno)d라인 - %(name)s - %(levelname)s - %(message)s')
    streamHandler = logging.StreamHandler()
    streamHandler.setFormatter(formatter)
    Logger.addHandler(streamHandler)
    fileHandler = logging.FileHandler(logFileName)
    Logger.addHandler(fileHandler)   
    fileHandler.setFormatter(formatter)

    return Logger

def getUserLogger(logFileName, name):
    Logger = logging.getLogger(name)
    Logger.setLevel(logging.INFO)
    streamHandler = logging.StreamHandler()
    Logger.addHandler(streamHandler)
    fileHandler = logging.FileHandler(logFileName)
    Logger.addHandler(fileHandler)   

    return Logger

def getYear():
    return "{}".format("%04d" % time.localtime().tm_year)

def getMonth():
    return "{}".format("%02d" % time.localtime().tm_mon)

def getDay():
    return "{}".format("%02d" % time.localtime().tm_mday)

# 정산작업에서 전달 시작일, 마지막일 반환
def getLastDate():
    today = date.today()
    last_month = today.replace(day=1) - timedelta(days=1)
    end = last_month.strftime("%Y.%m.%d")
    last_month = last_month.replace(day=1)
    start = last_month.strftime("%Y.%m.%d")
    return start, end

# 정산작업에서 익월 시작일, 15일 반환
def getLastDate():
    today = date.today()
    this_month = today.replace(day=1)
    start = this_month.strftime("%Y.%m.%d")
    this_month = today.replace(day=15)
    end = this_month.strftime("%Y.%m.%d")
    return start, end

# 브랜드명이 적힌 파일을 읽어와서 프로세스별로 일감분리
def divideWork(filename, partition):
    f = open(filename, 'r', encoding='CP949')
    rdr = csv.reader(f)
    result = []
    cnt = 1
    for i in range(0, partition):
        result.append([cnt])
        cnt += 1

    index = 0
    for item in rdr:
        if len(item) != 0:
            result[index % partition].append(item)
            index += 1
    f.close()
    return result

def mkdir(filename):
    try:
        if not os.path.isdir(filename):
            os.makedirs(os.path.join(filename))
    except OSError as e:
        if e.errno != errno.EEXIST:
            print("Failed to create " + filename + " directory!!!!!")

def exit():
    sys.exit()

# 실행 중 사용한 프로세스 정리
def taskkill():
    os.system('taskkill /f /im chrome.exe')
    os.system('taskkill /f /im excel.exe')
    os.system('taskkill /f /im chromedriver.exe')
    os.system('taskkill /f /im conhost.exe')
    os.system('taskkill /f /im python.exe')