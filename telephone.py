from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import Workbook
from selenium.common.exceptions import NoSuchElementException

from tkinter import *
import chromedriver_autoinstaller
import time

def handle_button_click():
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)   
    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome(options=chrome_options)

    myID=idEntry.get().replace(" ","")
    myPW=pwEntry.get().replace(" ","")
    telList=[]

    # 로그인창 접속
    driver.set_window_size(1200, 1000)  # 사이즈 조절: (가로, 세로)

    driver.get("https://www.knou.ac.kr")
    driver.find_element(By.XPATH, "//a[@id='btnLogin']").click()

    time.sleep(1)

    # 로그인
    userId = driver.find_element(By.ID, "username")
    userId.send_keys(myID)  # 로그인 할 계정 id
    userPwd = driver.find_element(By.ID, "password")
    userPwd.send_keys(myPW)  # 로그인 할 계정의 패스워드
    userPwd.send_keys(Keys.ENTER)

    time.sleep(1)

    # 학과튜터 로 계정 변경
    driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[@id='iframeContents']"))
    selectAccount = driver.find_element(By.XPATH, "//select[@id='chUser']")
    selectA=Select(selectAccount)
    selectA.select_by_index(2)
    driver.switch_to.parent_frame()

    time.sleep(2)

    # 강의 교수 지원 클릭
    ul_element = driver.find_element(By.XPATH, "//div[@class='gnbWrap']/ul[@class='menu']")
    li_element = ul_element.find_elements(By.TAG_NAME, "li")[1]
    a_element = li_element.find_element(By.TAG_NAME, "a")
    actions = ActionChains(driver)
    actions.move_to_element(a_element).perform()
    ul_element1= driver.find_element(By.XPATH, "//ul[@class='T2BG2']")
    li_element1 = ul_element1.find_element(By.TAG_NAME, "li")
    a_element1 = li_element1.find_element(By.TAG_NAME, "a")
    a_element1.click()

    time.sleep(1)

    # 학생분석보고서 및 배정조회 클릭
    driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[@id='iframeContents']"))
    ul_element2 = driver.find_elements(By.XPATH, "//ul[@class='mnb01']")[1]
    li_element2 = ul_element2.find_elements(By.TAG_NAME, "li")[4]
    a_element2 = li_element2.find_element(By.TAG_NAME, "a")
    a_element2.click()
    driver.switch_to.parent_frame()

    time.sleep(1)

    # 입학 예정자 선택
    driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[@id='iframeContents']"))
    selectState = driver.find_element(By.XPATH, "//select[@id='sregStDc']")
    selectS=Select(selectState)
    selectS.select_by_value("1")
    btn_search=driver.find_element(By.ID, "btn_search")
    driver.execute_script("arguments[0].click();",btn_search)

    time.sleep(1)
    now=1

    while (True):
        table=driver.find_element(By.XPATH, "//table[@class='ui-jqgrid-btable']")
        for i in range(1,11):
            strI=str(i)
            trElement=table.find_element(By.XPATH, f".//tr[@id='{strI}']")
            # 학생명
            td_name=trElement.find_element(By.XPATH, ".//td[@aria-describedby='jqgrid_studNm']")
            student_name=td_name.get_attribute("title")

            # 휴대폰 번호
            td_tel=trElement.find_element(By.XPATH, ".//td[@aria-describedby='jqgrid_indvTlno']")
            student_tel=td_tel.get_attribute("title")

            # 이메일
            td_mail=trElement.find_element(By.XPATH, ".//td[@aria-describedby='jqgrid_email']")
            student_mail=td_mail.get_attribute("title")

            # 연령대
            td_age=trElement.find_element(By.XPATH, ".//td[@aria-describedby='jqgrid_age']")
            student_age=td_age.get_attribute("title")

            semester=semEntry.get().replace(" ","")
            major=mjrEntry.get().replace(" ","")
            location=locEntry.get().replace(" ","")

            student_name_for_save=semester+" "+major+" "+location+" "+student_name+" ("+student_age+")"
            triplet=(student_name_for_save, student_tel, student_mail)
            telList.append(triplet)
        print(str(now)+"번째 페이지 저장 완료")
        now+=1
        try:
            ul_element3=driver.find_element(By.XPATH, "//ul[@class='pages']")
            li_element3=ul_element3.find_element(By.XPATH, ".//li[@class='pgNext pg-next']")
            driver.execute_script("arguments[0].click();",li_element3)
        except NoSuchElementException:
            print("페이지가 끝났습니다")
            break
        time.sleep(1)
        
    # 추출 정보를 엑셀 파일에 저장
    wb = Workbook()
    ws = wb.active
    column_names=["이름","전화번호", "이메일"]
    ws.append(column_names)
    for row, triplet in enumerate(telList, start=2):  # 첫 번째 행은 열 이름이므로 2부터 시작
        ws.append(triplet)
    wb.save("generic_test.xlsx")
    print("성공적으로 저장되었습니다!")

tk=Tk()
tk.title("튜터 연락처 저장 프로그램")
idLabel=Label(tk,text='ID  ').grid(row=0,column=0)
pwLabel=Label(tk,text='   PW').grid(row=0,column=2)
idEntry=Entry(tk)
pwEntry=Entry(tk)
idEntry.grid(row=0,column=1)
pwEntry.grid(row=0,column=3)

semLabel=Label(tk,text='학기  ').grid(row=1,column=0)
semEntry=Entry(tk)
semEntry.grid(row=1,column=1)
mjrLabel= Label(tk,text='   학과').grid(row=1,column=2)
mjrEntry=Entry(tk)
mjrEntry.grid(row=1,column=3)

locLabel=Label(tk,text='지역  ').grid(row=2,column=0)
locEntry=Entry(tk)
locEntry.grid(row=2,column=1)

infoLabel1=Label(tk,text='* 입력 공백 감지 X').grid(row=3,column=1)
infoLabel2=Label(tk,text='ex. 2024-1 사복 부산 홍길동 (20대)').grid(row=3,column=3)

btn1=Button(tk,text='저장 시작',bg='black',fg='white',command=handle_button_click).grid(row=4,column=2)
tk.mainloop()
