from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time

# ChromeDriver 경로 설정
chrome_service = Service('C:/Users/les04/chromedriver/chromedriver-win64/chromedriver.exe')
chrome_options = Options()

# 브라우저 시작
try:
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
except Exception as e:
    print(f"오류 발생: {e}")
    exit()

# 네이버 금융 페이지 접속
url = 'https://finance.naver.com/'
driver.get(url)

# 페이지 로드가 완료될 때까지 잠시 대기
time.sleep(3)

# 페이지 소스를 BeautifulSoup으로 파싱
soup = BeautifulSoup(driver.page_source, 'html.parser')

# 인기 종목 데이터 가져오기
tbody = soup.select_one('#container > div.aside > div.group_aside > div.aside_area.aside_popular > table > tbody')

# 데이터가 잘 로드되었는지 확인
datas = []
if tbody:
    trs = tbody.select('tr')
    for tr in trs:
        name = tr.select_one('th > a').get_text().strip()  # 종목명
        current_price = tr.select_one('td').get_text().strip()  # 현재가
        change_direction = tr['class'][0]  # 상승/하락 여부
        change_price = tr.select_one('td > span').get_text().strip()  # 변동가
        datas.append([name, current_price, change_direction, change_price])
else:
    print("데이터를 찾을 수 없습니다.")

# 브라우저 종료
driver.quit()

# 엑셀 파일 생성
wb = Workbook()
ws = wb.active
ws.title = '인기 종목'

# 헤더 추가
ws.append(['종목명', '현재가', '변동 방향', '변동가'])

# 데이터를 엑셀에 저장
for data in datas:
    ws.append(data)

# 엑셀 파일 저장 경로 (저장 경로는 본인의 환경에 맞게 수정)
file_path = 'C:/Users/les04/Documents/stock_data.xlsx'
wb.save(file_path)

print(f"데이터가 {file_path}에 저장되었습니다.")
