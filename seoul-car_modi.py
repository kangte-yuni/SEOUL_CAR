import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import datetime, time, sys, os
import openpyxl
import mysql.connector
from tqdm import tqdm
import urllib.request
import time
from PIL import Image
import os
import re
import pysftp
from selenium.webdriver.common.alert import Alert
#pip install pillow
#mysql 정보

# 추가
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui

mydb = None
"""
mydb = mysql.connector.connect(
    host = "127.0.0.1",
    port = "3306",
    database = "theday",
    user = "root", 
    password = "autoset",
    charset = 'utf8'
)

"""
mydb = mysql.connector.connect(
    host = "127.0.0.1",
    port = "3306",
    database = "theday",
    user = "root", 
    password = "0000",
    charset = 'utf8'
)

mycursor = mydb.cursor()

# 새로운 hidden 값 4개 체크 및 추가
table = 'g5_car_manager'
new_columns = ['car_maker','car_series','car_model','car_grade']

for column in new_columns:
    query = "SHOW COLUMNS FROM {} LIKE '{}'".format(table, column)
    mycursor.execute(query)

    if mycursor.fetchone():
        print("Column '{}' exists in table '{}'".format(column, table))
    else:
        print("Column '{}' does not exist in table '{}'".format(column, table))
        query = "ALTER TABLE {} ADD {} varchar(255);".format(table, column)
        mycursor.execute(query)
        print('Column {} added'.format(column))


#추출페이지수 세팅 ( 필요시 검색변수와 함께 인자로 받도록 처리 )
#page_ = input('최대 추출할 페이지 수를 입력하세요. (숫자만 입력) : ')
page_ = 3000
if page_ == '': page_ = 1

webdriver_options = webdriver.ChromeOptions()
#webdriver_options.add_argument('headless')
webdriver_options.add_argument("--window-size=1280,1080")
webdriver_options.add_experimental_option("excludeSwitches", ["enable-logging"])
webdriver_options.add_argument(f'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.87 Safari/537.36')

if getattr(sys, 'frozen', False): 
    # chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    # browser = webdriver.Chrome(chromedriver_path,options=webdriver_options)
    browser = webdriver.Chrome(ChromeDriverManager().install(), options=webdriver_options)
else:
    # browser = webdriver.Chrome(options=webdriver_options)
    browser = webdriver.Chrome(ChromeDriverManager().install(), options=webdriver_options)

# 쿠키 삭제하기
browser.delete_all_cookies()
    
login_url = 'http://carmanager.co.kr/User/Login'

browser.get(login_url)
#구성빈아이디
#browser.find_element_by_id('userid').send_keys('robotcc')
#browser.find_element_by_id('userpwd').send_keys('dustlr1004!')
#browser.find_element_by_xpath('//*[@id="ui_loginarea"]/tbody/tr/td[2]/button').click()
#박준섭 아이디
browser.find_element_by_id('userid').send_keys('jhgoodman7')
browser.find_element_by_id('userpwd').send_keys('wjdgns0153')
browser.find_element_by_xpath('//*[@id="ui_loginarea"]/tbody/tr/td[2]/button').click()
time.sleep(2)

############## 검색변수설정 (연동시 외부에서 넘겨줄 변수값) ################
# ex 경기,인천지역 / 전체 매물 / 10년이내 / 16만km 미만 / 가격체크 

s_area1 = ['서울'] #시도 (다중선택)
s_area2 = ['전체'] #지역 (다중선택 - 시도 다중선택시 선택불가)
s_area3 = ['',''] #단지 - area1,2 하나라도 다중선택시 선택불가 (ex 93011:한성단지, 93012:일산종합단지)

s_nation = '' #국산, 수입
s_maker = ''
s_model1 = ''
s_model2 = [''] #세부모델 (다중선택)
s_model3 = '' #multi R20
s_model4 = '' #multi 고급형

s_mission = ''
s_fuel = ''
s_color = '' #검정색

s_year1 = int(datetime.datetime.now().year)-10
s_month1 = datetime.datetime.now().month

s_year2 = ''
s_month2 = ''

s_mileage1 = ''
s_mileage2 = ''

s_price1 = '' #최저가격(만원)
s_price2 = '' #최고가격(만원)

s_extra3 = 1 #확장옵션 가격체크
###########################################################

browser.get('http://carmanager.co.kr/Car/Data')

wb = openpyxl.Workbook() # Workbook 생성 
sheet = wb.active # Sheet 활성 
# 헤더입력 
sheet.append(["고유번호", "차명", "차량번호", "미션", "연식", "연료", "주행", "색상", "가격", "지역(단지)", "담당/연락처", "사진정보","주요옵션","편의/안전옵션","영상/음향옵션","내/외장옵션","추가장착옵션","사고내역","압류건수","저당건수","세금미납정보","차량정보 상세내역","성능점검","최초등록일","수정일"]) 

# 체크박스, 라디오버튼 체크형식
def srh_check(keyword, xpath, type='check'):
    browser.find_element_by_xpath(xpath+'/button').click()
    if browser.find_element_by_xpath(xpath+'/div').get_attribute('style') == 'display: none;':
        browser.find_element_by_xpath(xpath+'/button').click()
    
    i = 0
    for li in browser.find_elements_by_xpath(xpath+'/div/ul/li'):
        i += 1
        try:
            x = xpath + '/div/ul/li['+str(i)+']/label/input'
            input = browser.find_element_by_xpath(x)
            #print(li.text.strip() + ' ' + str(i))
            if type == 'multi':
                if input and li.text.strip() in keyword:
                    # 미체크시 체크
                    if not input.is_selected(): 
                        li.click()
                else:
                    if input.is_selected(): li.click() # 기본체크된 경우 해제
            else:
                if input and li.text.strip() == keyword:
                    # 미체크시 체크
                    if type == 'check' and 'selected' not in li.get_attribute('class'): 
                        li.click()
                        break

        except:
            break # 선택된 값이 없습니다.
    
    browser.find_element_by_xpath('//*[@id="cs_searchcardetail"]/h3').click() #검색선택 레이어닫기

# 단순 a링크 셀렉트형식
def srh_select(keyword, xpath):
    browser.find_element_by_xpath(xpath+'/a[2]').click()
    for a in browser.find_elements_by_xpath(xpath+'/ul/li/a'):
        if a.text.strip() == str(keyword):
            a.click()
            break    

# 지역1 (시도) : 전체, 서울, 경기, 인천, 강원, 대전, 충남 ...
if s_area1:
    xpath = '//*[@id="ui_search"]/div/div[1]/div/div[1]'
    srh_check(s_area1, xpath, 'multi')

# 지역2 (구군)
if s_area2:
    xpath = '//*[@id="ui_search"]/div/div[1]/div/div[2]'
    srh_check(s_area2, xpath, 'multi')

# 지역3 (단지)
if s_area3:
    xpath = '//*[@id="ui_search"]/div/div[1]/div/div[3]'
    srh_check(s_area3, xpath, 'multi')

# 구분 (국산/수입)
if s_nation:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[1]/div[1]'
    srh_check(s_nation, xpath)

# 제조사
if s_maker:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[1]/div[2]'
    srh_check(s_maker, xpath)

# 모델 (그랜저)
if s_model1:
    xpath = '//*[@id="searchModel_multi_div"]/div'
    srh_check(s_model1, xpath)

# 세부모델 (뉴그랜저XG)
if s_model2:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[1]/div[4]'
    srh_check(s_model2, xpath, 'multi')

# 모델등급 (R20)
if s_model3:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[1]/div[5]'
    srh_check(s_model3, xpath, 'multi')

# 세부등급 (전체/기본형/고급형)
if s_model4:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[1]/div[6]'
    srh_check(s_model4, xpath, 'multi')

# 미션
if s_mission:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[2]/div[1]'
    srh_select(s_mission, xpath)

# 연료
if s_fuel:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[2]/div[2]'
    srh_select(s_fuel, xpath)

# 색상
if s_color:
    xpath = '//*[@id="cs_searchcardetail"]/div/div[2]/div[3]'
    srh_select(s_color, xpath)

# 연식
#if s_year1:
#    xpath = '//*[@id="cs_searchcardetail"]/div/div[2]/div[4]/div[1]'
#    srh_select(s_year1, xpath)

#if s_month1:
#    xpath = '//*[@id="cs_searchcardetail"]/div/div[2]/div[4]/div[2]'
#    srh_select(s_month1, xpath)

#if s_year2:
#    xpath = '//*[@id="cs_searchcardetail"]/div/div[2]/div[4]/div[3]'
#    srh_select(s_year2, xpath)

#if s_month2:
#    xpath = '//*[@id="cs_searchcardetail"]/div/div[2]/div[4]/div[4]'
#    srh_select(s_month2, xpath)

# 주행거리
if s_mileage1 or s_mileage2:
    browser.find_element_by_id('cbxSearchKmCheck').click()
    if s_mileage1 : browser.find_element_by_id('tbxSearchDriveSI').send_keys(s_mileage1)
    if s_mileage2 : browser.find_element_by_id('tbxSearchDriveEI').send_keys(s_mileage2)

# 가격
if s_price1 or s_price2:
    browser.find_element_by_id('cbxSearchMoneyCheck').click()
    if s_price1 : browser.find_element_by_id('tbxSearchMoneySI').send_keys(s_price1)
    if s_price2 : browser.find_element_by_id('tbxSearchMoneyEI').send_keys(s_price2)

if s_extra3: # 확장검색 가격체크, 사진체크
    browser.find_element_by_id('cbxSearchIsCarPhoto').click()
    browser.find_element_by_id('cbxSearchIsSaleAmount').click()

#차량번호 입력
    #browser.find_element_by_id('tbxSearchCarNumber').send_keys('171서9984')
    
# 검색버튼실행
browser.find_element_by_xpath('//*[@id="ui_search"]/div/div[5]/button[2]').click()
time.sleep(3)

# 50개씩 보기
# srh_select('50', '//*[@id="ui_isphotocol"]/li[8]/div')
# time.sleep(2)

# 검색데이터 여부
if browser.find_element_by_xpath('//*[@id="ui_context"]/table/tbody/tr[2]/td').text == '데이터가 없습니다.':
    print('검색데이터 없음 - 종료')
    exit()

# 마지막페이지 추출
lastPage_btn = browser.find_element_by_xpath('//*[@id="ui_cfooter"]/div[1]/div[2]/div/table/tbody/tr/th[4]').get_attribute('onclick')
if lastPage_btn:
    lastPage = lastPage_btn.replace("goPageSubmit(","").replace(")","")
else:
    lastPage = 1


#입력받은 페이지보다 큰경우 입력페이지로 제한 
if int(lastPage) > int(page_):
    lastPage = page_


#중복체크를위해 카우트 0으로 처리

sql_reset = " update g5_car_manager set car_cnt = 1, car_cnt_date = now() where car_area_type = %s  "
val_reset = ('서울',)                  
mycursor.execute(sql_reset, val_reset)
mydb.commit()

start_time = datetime.datetime.now()
print(start_time)
# 결과데이터 추출정리 #################################################################
#중복값체크
other = 1
others = 1
country = ''
for p in range(1, int(lastPage)+1):
    print('['+ str(p) + ' 페이지]')
    #기본 1 페이지는 출력
    listTr = browser.find_elements_by_xpath('//*[@id="ui_context"]/table/tbody/tr')
    
    if len(listTr) > 2:
        for tr in tqdm(listTr[1:-1]):
            
            data = []
            try:
                idx = tr.find_element_by_xpath('.//td[1]/input').get_attribute('value')
                a = tr.find_element_by_xpath('.//td[2]/a')

                browser.execute_script("window.scrollTo(0, window.pageYOffset+78);")

                thumb = tr.find_element_by_xpath('.//td[2]/a/img').get_attribute('src')
                title = tr.find_element_by_xpath('.//td[4]').text.strip() #상세페이지에서 추출
                titles = title.split('\n')
                title = titles[0];

                pqs = re.compile(r"\[(.+)\](.+)")
                m = pqs.match(title)
                pname = str(m.group(1))

                if pname == "현대" or pname == "삼성" or pname == "대우" or pname == "쉐보레" or pname == "기아" or pname == "쌍용" or pname == "제네시스" or pname == "중대형화물" or pname == "중대형버스" or pname == "기타제조사" :
                    country = "한국"
                else :
                    country = "외국"

                mission = tr.find_element_by_xpath('.//td[5]').text.strip()
                birth = tr.find_element_by_xpath('.//td[6]').text.strip()
                fuel = tr.find_element_by_xpath('.//td[7]').text.strip()
                #mileage = tr.find_element_by_xpath('.//td[8]').text.strip()
                color = tr.find_element_by_xpath('.//td[9]').text.strip()
                
                price = tr.find_element_by_xpath('.//td[10]').text.strip()
                price = price.replace(',', '')
                
                area = tr.find_element_by_xpath('.//td[11]').text.strip()
                manager = tr.find_element_by_xpath('.//td[12]').text.strip()

                #중복체크로직 추가합니다.
                sql_j = " select car_money,car_images from g5_car_manager where car_code = %s "
                val_dc = (idx,)
                mycursor.execute(sql_j, val_dc)
                rows = mycursor.fetchall()

                
                for dataj in rows:
                    print("중복 : " + idx)
                    #print(dataj[0])
                    #print(price)
                    if dataj[0] != price:
                        #가격만 업데이트
                        print("가격변경")
                        sql_uu = " update g5_car_manager set car_money = %s, car_money_update = now() where car_code = %s "
                        val_dcu = (price, idx)
                        mycursor.execute(sql_uu, val_dcu)
                        mydb.commit()

                    #판매되있는 확인하는 cnt
                    print("차량있음")
                    sql_cnt = " update g5_car_manager set car_cnt = car_cnt + 1, car_db_updatedate = now() where car_code = %s "
                    val_cnt = (idx,)
                    mycursor.execute(sql_cnt, val_cnt)
                    mydb.commit()

             
                    if dataj[1] == '':
                        other = 1
                        #추가작업할경우
                        others = 2
                    else :
                        other = 2         
                        
                
                if other == 1:
                    
                    # 상세페이지 클릭(새창)
                    a.click()
                    browser.switch_to.window(browser.window_handles[1])
                    time.sleep(3)
                
            except:
                break
                
                

            if other == 1:
		# alert창 대신 enter키 입력을통한 오류 해결
                try:
                    # da = Alert(browser)
                    # da.accept()
                    pyautogui.FAILSAFE = False
                    pyautogui.hotkey('enter')
                    alert = browser.switch_to.alert
                    alert.accept()
                    alert.dismiss()

                except:
                    pyautogui.FAILSAFE = False
                    pyautogui.hotkey('enter')
                    #alert = browser.switch_to.alert
                    #alert.accept()
                    #alert.dismiss()
                    pass
                #사진정보 (기본탭)
                img_urls = []
                img_urls_ftp = []
                bigimgs = browser.find_elements_by_xpath('//*[@id="photoPage1"]/li/a/img[2]')
                imgcnt = int(1)

                #ftp 접속정보
                host = '@@' # 호스트명만 입력. sftp:// 는 필요하지 않다.
                port = 22 # int값으로 sftp서버의 포트 번호를 입력
                username = '@@' # 서버 유저명
                password = '@@' # 유저 비밀번호

                hostkeys = None

                # 서버에 저장되어 있는 모든 호스트키 정보를 불러오는 코드
                cnopts = pysftp.CnOpts()

                # 접속을 시도하는 호스트에 대한 호스트키 정보가 존재하는지 확인
                # 존재하지 않으면 cnopts.hostkeys를 None으로 설정해줌으로써 첫 접속을 가능하게 함

                if cnopts.hostkeys.lookup(host) == None:
                    print("Hostkey for " + host + " doesn't exist")
                    hostkeys = cnopts.hostkeys # 혹시 모르니 다른 호스트키 정보들 백업
                    cnopts.hostkeys = None

                try :

                    with pysftp.Connection(
                            host,
                            port = port,
                            username = username,
                            password = password,
                            cnopts = cnopts) as sftp:

                            sftp.chdir("/home/carimg/imgs2")
                            print(sftp.pwd)
                            print(idx)
                            idxs = str(idx)
                            print(sftp.isdir(str(idxs)))

                            if sftp.isdir(str(idxs)) == False :
                                sftp.mkdir(idxs)
                                #파일전송
                                #sftp.put(filename, preserve_mtime=True)
                                

                    
                            for bimg in bigimgs:
                                bimg_src = bimg.get_attribute('src')
                                bimg_src = bimg_src.replace("?watermark=rb", "")
                                img_urls.append(bimg_src)
                                bimg_src_JPG = bimg_src.upper().split('JPG')
                                bimg_src = bimg_src_JPG[0] + "jpg"                           
                                
                                #줄바꿈으로 이미지 가져오기
                                #x = img_data.splitlines()

                                #print("첫 번쨰 : ", x[0])
                                #print("두 번쨰 : ", x[1])
                                
                                # 다운받을 이미지 url
                                url = bimg_src
                                # time check
                                start = time.time()
                                # 이미지 요청 및 다운로드

                                directory = "d:\\py\\imgs2\\"+idx
                                #print(directory)
                                try:
                                    if not os.path.exists(directory):
                                        print("생성")
                                        os.makedirs(directory)
                                except OSError:
                                    print("Error: Failed to create the directory.")

                        
                                try:
                                    filenames = "D:\\py\\imgs2\\"+idx+"\\"+idx+"_"+str(imgcnt)+".jpg"
                                    filenames1 = ""+idx+"_"+str(imgcnt)+".jpg"
                                    urllib.request.urlretrieve(url, filenames)
                                except:
                                    pass
                                print(filenames1);
                                # 이미지 다운로드 시간 체크
                                print(time.time() - start)
                                # 저장 된 이미지 확인
                                #img = Image.open("test.jpg")
                                imgcnt += int(1)

                               
                              
                                sftp.chdir("/home/carimg/imgs2/"+idx)
                                #print(sftp.pwd)
                                try :
                                    sftp.put(filenames, preserve_mtime=True)
                                except:
                                    pass

                                img_urls_ftp.append(filenames1)               
                except:
                    pass
        
                img_data = ('\n').join(img_urls)
                img_data_ftp = ('\n').join(img_urls_ftp)

      
                #ftp.close

                #차명
                #title = browser.find_element_by_xpath('//*[@id="content"]/div[5]/table/tbody/tr[1]/td[1]').text.strip()
                
                #차량번호
                carNumber = browser.find_element_by_xpath('//*[@id="carplatenoCopy"]').get_attribute('value')

                #km
                mileage = browser.find_element_by_xpath('//*[@id="content"]/div[5]/table/tbody/tr[3]/td[1]/span[1]').text.strip()
                mileage = mileage.replace(",", "")
                
                #제시번호
                jesi = browser.find_element_by_xpath('//*[@id="content"]/div[5]/table/tbody/tr[3]/td[5]').text.strip()
                
                #최초등록일
                redDate = browser.find_element_by_xpath('//*[@id="ui_ViewCarReg"]').text.strip()

                #수정일
                modDate = browser.find_element_by_xpath('//*[@id="content"]/div[2]/div[2]/span[1]').text.replace('수정일 : ','').strip()
                
                #상세옵션탭 이동
                browser.find_element_by_xpath('//*[@id="content"]/div[2]/ul/li[2]/a').click()
                
                #주요옵션 체크사항
                optArr1 = []
                opt1s = browser.find_elements_by_xpath('//*[@id="option_ul"]/li')
                for opt1 in opt1s:
                    input = opt1.find_element_by_tag_name("input")
                    label = opt1.find_element_by_tag_name("label").text.strip()
                    if input.is_selected():
                        optArr1.append(label)
                opt1_data1 = ('\n').join(optArr1)

                #편의안전 체크사항
                optArr2 = []
                opt1s = browser.find_elements_by_xpath('//*[@id="safety_ul"]/li')
                for opt1 in opt1s:
                    input = opt1.find_element_by_tag_name("input")
                    label = opt1.find_element_by_tag_name("label").text.strip()
                    if input.is_selected():
                        optArr1.append(label)
                opt1_data2 = ('\n').join(optArr2)

                #영상음향 체크사항
                optArr3 = []
                opt1s = browser.find_elements_by_xpath('//*[@id="videoacoustic_ul"]/li')
                for opt1 in opt1s:
                    input = opt1.find_element_by_tag_name("input")
                    label = opt1.find_element_by_tag_name("label").text.strip()
                    if input.is_selected():
                        optArr1.append(label)
                opt1_data3 = ('\n').join(optArr3)

                #내외장 체크사항
                optArr4 = []
                opt1s = browser.find_elements_by_xpath('//*[@id="etc_ul"]/li')
                for opt1 in opt1s:
                    input = opt1.find_element_by_tag_name("input")
                    label = opt1.find_element_by_tag_name("label").text.strip()
                    if input.is_selected():
                        optArr1.append(label)
                opt1_data4 = ('\n').join(optArr4)
                
                #추가장착옵션
                addopt1 = browser.find_element_by_xpath('//*[@id="ui_popup_cardetail"]/div/div[2]/div/dl/dd[1]').text.strip()

                #사고내역
                addopt2 = browser.find_element_by_xpath('//*[@id="ui_popup_cardetail"]/div/div[2]/div/dl/dd[2]').text.strip()

                #압류건수
                addopt3 = browser.find_element_by_xpath('//*[@id="ui_popup_cardetail"]/div/div[2]/div/dl/dd[3]').text.strip()

                #저당건수
                addopt4 = browser.find_element_by_xpath('//*[@id="ui_popup_cardetail"]/div/div[2]/div/dl/dd[4]').text.strip()

                #세금미납정보
                addopt5 = browser.find_element_by_xpath('//*[@id="ui_popup_cardetail"]/div/div[2]/div/dl/dd[5]').text.strip()

                #차량정보상세
                addopt6 = browser.find_element_by_xpath('//*[@id="ui_popup_cardetail"]/div/div[2]/div/dl/dd[6]/textarea').text.strip()
                
                #차량 제조사
                carMaker = browser.find_element_by_xpath('//*[@id="CAR_MAKER"]').get_attribute('value')
                
                #차량 이름
                carSeries = browser.find_element_by_xpath('//*[@id="CAR_NAME"]').get_attribute('value')
                
                #차량 모델명
                carModel = browser.find_element_by_xpath('//*[@id="CAR_MODEL"]').get_attribute('value')
                
                #차량 등급
                carGrade = browser.find_element_by_xpath('//*[@id="CAR_GRADE"]').get_attribute('value')
                
                
                #성능점검탭 이동
                browser.find_element_by_xpath('//*[@id="content"]/div[2]/ul/li[3]/a').click()
                
                try:
                    # da = Alert(browser)
                    # da.accept()

                    pyautogui.FAILSAFE = False
                    pyautogui.hotkey('enter')
                    alert = browser.switch_to.alert
                    alert.accept()
                    alert.dismiss()
                except:
		    #alert = browser.switch_to.alert
                    #alert.accept()
                    #alert.dismiss()
                    pyautogui.FAILSAFE = False
                    pyautogui.hotkey('enter')
                    pass
                
                carcheckoutUrl = browser.find_element_by_id('carcheckout').get_attribute('src')
                
                browser.close() #새창닫고
                browser.switch_to.window(browser.window_handles[0]) #기존창으로 복귀


                #db insert

                if others == 1:
                    sql = "INSERT INTO g5_car_manager (car_code, car_country, car_name, car_number, car_auto, car_yymm, car_gas, car_km, car_jesi, car_color, car_money, car_area, car_manager, car_opt0, car_opt1, car_opt2, car_opt3, car_opt4, car_opt5, car_url, car_images, car_images_ori, car_insert_date, car_db_insertdate, car_cnt, car_area_type, car_maker, car_series, car_model, car_grade) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, now(), 2, '서울', %s, %s, %s, %s);"
                    val = (idx, country, title, carNumber, mission, birth, fuel, mileage, jesi, color, price, area, manager, opt1_data1, addopt2, addopt3, addopt4, addopt5, addopt6, carcheckoutUrl, img_data_ftp, img_data, redDate, carMaker, carSeries, carModel, carGrade)
                elif others == 2:
                    dsql = " DELETE from g5_car_manager where car_code = %s "
                    dval_dcu = (idx,)
                    mycursor.execute(dsql, dval_dcu)
                    mydb.commit()
                    
                    sql = "INSERT INTO g5_car_manager (car_code, car_country, car_name, car_number, car_auto, car_yymm, car_gas, car_km, car_jesi, car_color, car_money, car_area, car_manager, car_opt0, car_opt1, car_opt2, car_opt3, car_opt4, car_opt5, car_url, car_images, car_images_ori, car_insert_date, car_db_insertdate, car_cnt, car_area_type, car_maker, car_series, car_model, car_grade) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, now(), 2, '서울', %s, %s, %s, %s);"
                    val = (idx, country, title, carNumber, mission, birth, fuel, mileage, jesi, color, price, area, manager, opt1_data1, addopt2, addopt3, addopt4, addopt5, addopt6, carcheckoutUrl, img_data_ftp, img_data, redDate, carMaker, carSeries, carModel, carGrade)
                #sql = "INSERT IGNORE INTO g5_car_manager (car_code, car_name, car_db_insertdate) VALUES (%s, %s, now());"
                #val = (idx, title)          
                try :
                    mycursor.execute(sql, val)
                    mydb.commit()
                except:
                    pass




                #sheet.append([idx,title,carNumber,mission,birth,fuel,mileage,color,price,area,manager,img_data,opt1_data1,opt1_data2,opt1_data3,opt1_data4,addopt1,addopt2,addopt3,addopt4,addopt5,addopt6,carcheckoutUrl,redDate,modDate])
                #print(idx)
            other = 1
            others = 1
    print(datetime.datetime.now())
    # 페이징버튼 보이도록 최하단스크롤 후 클릭
    #browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    pindex = ( p % 5 ) + 1

    if int(p) == int(lastPage):
       break
    elif pindex == 1 and int(p) < int(lastPage):
        #browser.find_element_by_xpath('//*[@id="ui_cfooter"]/div[1]/div[2]/div/table/tbody/tr/th[3]').click() #다음블럭페이징 클릭
        #browser.find_element_by_xpath('/html/body/div[8]/div[2]/form/section/div/div/div/div[3]/div[1]/div[2]/div/table/tbody/tr/th[3]').send_keys(Keys.ENTER)
        element2 = browser.find_element_by_xpath('//*[@id="ui_cfooter"]/div[1]/div[2]/div/table/tbody/tr/th[3]')
        browser.execute_script("arguments[0].click();", element2)
    else:
        #browser.find_element_by_xpath('//*[@id="ui_cfooter"]/div[1]/div[2]/div/table/tbody/tr/td['+str(pindex)+']').click()
        #browser.find_element_by_xpath('/html/body/div[8]/div[2]/form/section/div/div/div/div[3]/div[1]/div[2]/div/table/tbody/tr/td['+str(pindex)+']').send_keys(Keys.ENTER)
        element1 = browser.find_element_by_xpath('//*[@id="ui_cfooter"]/div[1]/div[2]/div/table/tbody/tr/td['+str(pindex)+']')
        browser.execute_script("arguments[0].click();", element1)
        
    time.sleep(8)


browser.close()
print(start_time)
print(datetime.datetime.now())
print('======= 추출완료!! =======')
wb.save('carmanager.xlsx')
wb.close()