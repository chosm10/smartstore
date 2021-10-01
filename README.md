# 2020.12.21 빗버킷 이관 완료
# Naver Smart Store Crawling
## 파이썬 버젼 3.X 설치
- Windows는 C:/Users/Administrator/Desktop/chromedriver_win32/에, Linux는 /naver/chromedriver/에 운영체제별 chromedriver 설치 필요!
## 실행 방법
- git에서 설치한 디렉터리에서 python 업무명(ex - naverdaily.py) 점포명(ex - ddp, gimpo) / 동일 명령어가 적힌 .bat 실행    
```python naverdaily.py ddp```
## pip 설치가 안될 시
- get-pip.py를 실행하여 pip 수동 설치
## 모듈(라이브러리) 종속성 관리
- requirements.txt에 설치 모듈 버젼이 기록되어 있음. 다음 명령어 실행(-r 옵션: 포맷파일을 읽어 설치)    
```pip install -r requirements.txt```
## xpath 및 시스템 관련 정보 관리
- data.json 파일에서 일괄 관리, json 객체로 읽어서 프로그램 상 활용
## 크롬드라이버 관리
- 프로그램을 실행할 PC의 버젼(현재 chrome version: 85.0.4183.102) 에 맞는 크롬드라이버 설치 필요
- 설치 경로: C:\Users\Administrator\Desktop\chromedriver_win32  (필요에 따라 경로 변경 가능, data.json의 driver에서 경로 재설정) 
## 점포별 스토어 리스트 관리
- brands_점포명.csv 파일에 저장 
## 활용 멀티프로세싱 수 변경
- 각 업무 메인 프로그램의 __main__에서 변경 가능, 테스트 결과 4코어 이상 사용 시 네이버 스마트 스토어에서 로봇으로 인지하여 __홈페이지 사용 불가__
## 업무별 메인 프로그램
- 일매출정리: naverdaily.py
## 공통 기능 모듈
- naver.py: 네이버 스마트스토어 관련 공통 부분 기능 제공
- api.py: 로그, 캡쳐, 시간 등의 기능 제공
- mail.py: 메일 발송 기능 제공
- excel_concat.py: 엑셀 병합 기능 제공
## 점포별 스토어 리스트 파일 생성 프로그램
- storeCrawl.py 실행
## 크롬 드라이버 브라우져 버젼에 상관 없이 작동하게 하는 부분
- 크롬 드라이버 매니져를 사용해서 실행 당시의 드라이버를 다운해서 사용
```driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=option)```
## 네이버 로그인 창 다른 점포 처리
- naver.py의 loginSelect 변수로 로그인 화면 구분하여 분기 처리됨
## 점포 추가 순서
- brands_점포명.csv 생성 -> data.json 파일의 members에 계정 정보 추가 and workFileName에 해당 점포가 사용할 브랜드리스트명 추가 ->   
- ```python brandList.py 점포명``` 으로 brands.csv에 해당 점포의 브랜드 리스트 받기 -> brands.csv 내용을 해당 점포 브랜드 리스트에 복사 / 붙여넣기 
## 네이버 정산 프로그램에서 기존 날짜 칸 설정 방법 변경
- 기존에는 날짜칸 인식 -> readonly 해제 -> sendkey.enter로 초기화 -> sendkey(정해진날짜) 였음
- sendkey 방식은 팝업창 같은게 발생해서 화면이 가려지면 인식이 안되는 경우 발생
- 자바스크립트 조작 방식으로 setAttribute로 value 속성값을 직접 변경: 날짜칸 인식 -> setAttribute('value', 정해진날짜)