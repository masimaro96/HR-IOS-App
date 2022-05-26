from asyncio import events
from distutils.log import Log
import time, sys, unittest, random, json, openpyxl, platform
from datetime import datetime
from appium import webdriver
from random import randint
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Color, PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles import Color, Fill
from openpyxl.cell import Cell

from NQ_function import Logging, data
from framework_sample import *

APPIUM_PORT = '4723'
udid = 'bc86e429485c13f34837866fde36e7ed55646317'
app_path = 'Users/hanbiro/Desktop/nhuquynhios/HanbiroHR.ipa'
command_executor ='http://127.0.0.1:%s/wd/hub' % APPIUM_PORT

'''desired_capabilities = {
    'orientation' :'LANDSCAPE',
    "deviceName": "Hanbiro Iphone",
    "platformVersion": "12.5.5",
    "platformName": "IOS",
    "udid": udid,
    "app": app_path
}'''

desired_capabilities = {
    #"xcodeOrgId": "9689HPSFXL",
    #"xcodeSigningId": "iPhone Developer",
    "deviceName": "Hanbiro Iphone",
    "platformName": "IOS",
    "orientation": "PORTRAIT",
    #"newCommandTimeout": 0,
    "automationName": "XCUITest",
    "derivedDataPath" : "/Users/hanbiro/Library/Developer/Xcode/DerivedData/WebDriverAgent-aghlrsejdreqngftgvcqwnjgrbou",
    "wdaConnectionTimeout": 500000,
    "udid": udid,
    "app": app_path
}

driver = webdriver.Remote(command_executor, desired_capabilities)

n = random.randint(1,3000)


def execution():
    Logging("------- Login to app -------")
    # Input information for log-in
    driver.find_element_by_ios_class_chain(data["domain"]).send_keys(data["domain_name"])
    Logging("- Input Domain")
    driver.find_element_by_ios_class_chain(data["id_login"]).send_keys(data["id_login_name_1"])
    Logging("- Input ID")
    driver.find_element_by_ios_class_chain(data["pass_login"]).send_keys(data["password"])
    Logging("- Input Password")
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["login_button"]))).click()
    Logging("=> Click Log In button")
    driver.implicitly_wait(1000)

    add_event()


def clock_in():
    try:
        Logging("--- Clock in with GPS ---")
        try:
            OT = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["OT"]["status_OT"])))        
            if OT.text == 'Night shift':
                Logging("=> Work night shift")
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["OT"]["confirm_OT"]))).click()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["OT"]["apply_OT"]))).click()
                Logging("=> Confirm OT")

                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["OT"]["reason_OT"]))).send_keys(data["OT"]["reason_text"])
                Logging("=> Input reason OT")

                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["OT"]["confirm_apply_OT"]))).click()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["OT"]["button_close"]))).click()
                Logging("=> Apply OT success")
            else:
                Logging("=> Apply OT not display")
        except WebDriverException:
            Logging("=> Apply OT not display")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_in"]["button_clock_in"]))).click()
        Logging("=> Click clock in")

        status_clock_in= WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_in"]["status_late"])))
        if status_clock_in.text == 'Tardiness':
            Logging("=> Clock in late")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_in"]["reason_late"]))).send_keys(data["clock_in"]["reason_late_text"])
            Logging("=> Input reason late")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_in"]["button_save"]))).click()
            Logging("=> Save")
        else:
            Logging("=> Clock in on time")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_in"]["button_close"]))).click()

    except WebDriverException:
        Logging("=> Crash app")

def break_time():
    try:
        Logging("--- Break time ---")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["break_time"]["button_break_time"]))).click()
        Logging("=> Start break time")
        time.sleep(30)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["break_time"]["end_break_time"]))).click()
        Logging("=> End Breake time")
        time.sleep(10)
    except WebDriverException:
        Logging("=> User have clock out")

def clock_out():
    try:
        Logging("--- Clock out ---")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_out"]["button_clock_out"]))).click()
        Logging("=> Click clock out")

        status = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_out"]["status_clock_out"])))
        if status.text == 'Leave Early':
            Logging("=> Clock out early")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_out"]["reason_clock_out"]))).send_keys(data["clock_out"]["reason_clock_out_text"])
            Logging("=> Input reason")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_out"]["confirm_clock_out"]))).click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_out"]["clock_out_success"]))).click()
            Logging("=> Clock out success")
        else:
            Logging("=> Clock out on time")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["clock_out"]["button_close"]))).click()
    except WebDriverException:
        Logging("=> Crash app")

def view_calendar():
    Wait10s_ClickElement(data["calender_home"]["date"])
    Logging("- Select calender")
    Wait10s_ClickElement(data["calender_home"]["date_to_view"])
    Wait10s_ClickElement(data["calender_home"]["select"])
    Logging("- Select date")

    popup = WaitElementLoaded(20, data["calender_home"]["popup_warning"])
    if popup == "Warning":
        Logging("- Date select not worked day -> Select again")
    else:
        Logging("- Date correct")

    Wait10s_ClickElement(data["calender_home"]["preview_date"])
    Logging("- View preview date")
    Wait10s_ClickElement(data["calender_home"]["next_date"])
    Logging("- View next date")

def view_noti():
    Logging("--- View Notification ---")
    Wait10s_ClickElement(data["menu_settings"]["button_notification"])
    Logging("- View content notification")
    Wait10s_ClickElement(data["menu_settings"]["button_back"])
    Logging("-> Back to menu")

def settings():
    try:
        Logging("--- Settings - Change language ---")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='한국어']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='닫기']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()
        
        korean_text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["settings"]["language_text"])))
        if korean_text.text == '한국어':
            Logging("=> Change to language '한국어' success")
        else:
            Logging("=> Fail")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='언어 설정']"))).click()
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Tiếng Việt']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Đóng']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()

        VN_text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["settings"]["language_text"])))
        if VN_text.text == 'Tiếng Việt':
            Logging("=> Change to language 'Tiếng Việt' success")
        else:
            Logging("=> Fail")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Thay đổi ngôn ngữ']"))).click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='日本語']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='閉じる']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()

        JP = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["settings"]["language_text"])))
        if JP.text == '日本語':
            Logging("=> Change to language '日本語' success")
        else:
            Logging("=> Fail")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='言語']"))).click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='简体中文']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='關閉']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()

        TQ = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["settings"]["language_text"])))
        if TQ.text == '简体中文':
            Logging("=> Change to language '简体中文' success")
        else:
            Logging("=> Fail")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='語言']"))).click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Indonesian']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Tutup']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()

        indo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["settings"]["language_text"])))
        if indo.text == 'Indonesian':
            Logging("=> Change to language 'Indonesian' success")
        else:
            Logging("=> Fail")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Ganti BAHASA']"))).click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='English']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Close']"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()
        EN = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["settings"]["language_text"])))
        if EN.text == 'English':
            Logging("=> Change to language 'English' success")
        else:
            Logging("=> Fail")
        
    except WebDriverException:
        Logging("=> Crash app")

def add_event():
    try:
        Logging(" ")
        Logging("------- Add event -------")
        Wait10s_ClickElement(data["menu_timecard"]["button_timecard"])
        Logging("- Select time card")
        Wait10s_ClickElement(data["menu_timecard"]["button_timesheet"])
        Logging("- Select time sheet")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["add_event"]))).click()
        Logging("- Select add")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@text,'Please input data.')]"))).send_keys(data["event"]["title_text"])
        Logging("- Input title")
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["choose_event"]))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["type_event"]))).click()
        Logging("- Choose event")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["select_color"]))).click()
        Logging("- Select color")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["place"]))).send_keys(data["event"]["location_text"])
        Logging("- Input location")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["memo"]))).send_keys(data["event"]["memo_text"])
        Logging("- Input memo")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["button_save"]))).click()
    except:
        Logging("- Can't create event")

    Logging("** Check event use approval type")
    try:
        approval_type = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["popup"])))
        if approval_type.text == '[Approved] Your request approval request has been approved automatically':
            Logging("=> Use approval type: Automatic approval")

        elif approval_type.text == 'The approval request has been submitted. Please wait until the approval is completed.':
            Logging("=> Use approval type: Approval Line")

        elif approval_type.text == 'The approval request has been delivered to Head of Department. Please wait until the approval is completed.':
            Logging("=> Use approval type: Head Dept.")

        elif approval_type.text == 'The approval request has been delivered to Timecard Managers. Please wait until the approval is completed.':
            Logging("=> Use approval type: Timecard Manager")
        else:
            Logging("=> Use approval type: Dept. Manager")
    except WebDriverException:
        Logging("=> Use approval type: Dept. Manager") 

    driver.find_element_by_xpath(data["event"]["close_popup"]).click()
    Logging("=> Save event")
    time.sleep(5)

def request_vacation():
    try:
        Logging("-- Request vacation--")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_timecard"]["button_timecard"]))).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["button_vacation"]))).click()
        Logging("- Request Vacation")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_timecard"]["request_vacation"]))).click()
        Logging("Save request")
    except:
        Logging("-> Can't request vacation")

def admin_settings_GPS():
    Wait10s_ClickElement(data["menu_admin"]["button_admin"])
    Logging("- Admin settings")

    Wait10s_ClickElement(data["menu_admin"]["GPS_setting"])
    Logging("- GPS Settings")
    Wait10s_ClickElement(data["menu_admin"]["add_GPS"])
    Wait10s_ClickElement(data["menu_admin"]["popup"])
    Logging("- Add GPS")
    Wait10s_ClickElement(data["menu_admin"]["search_gps"])
    Logging("- Input GPS")
    InputElement(data["menu_admin"]["search_gps"], "Nguyen")
    Logging("- Enter value")
    Wait10s_ClickElement(data["menu_admin"]["done"])
    Wait10s_ClickElement(data["menu_admin"]["gps"])
    Logging("- Select GPS")
    Wait10s_ClickElement(data["menu_admin"]["workplace"])
    Wait10s_ClickElement(data["menu_admin"]["select_workplace"])
    Logging("- Select workplace")
    Wait10s_ClickElement(data["menu_admin"]["save_GPS"])
    Wait10s_ClickElement(data["menu_admin"]["close_popup"])
    Logging("- Save GPS")

    driver.swipe(start_x=1000, start_y=450, end_x=500, end_y=450, duration=800)
    Wait10s_ClickElement(data["menu_admin"]["delete_gps"])
    Wait10s_ClickElement(data["menu_admin"]["button_yes"])
    Logging("- Delete GPS")
    Wait10s_ClickElement(data["menu_admin"]["close_popup"])

def admin_settings_wifi():
    Wait10s_ClickElement(data["menu_admin"]["back_button"])
    Wait10s_ClickElement(data["menu_admin"]["Wifi_setting"])
    Logging("- Wifi Settings")
    Wait10s_ClickElement(data["menu_admin"]["add_wifi"])
    Logging("- Add Wifi")

def admin_settings_beacon():
    Wait10s_ClickElement(data["menu_admin"]["back_button"])
    Wait10s_ClickElement(data["menu_admin"]["Beacon_setting"])
    Logging("- Beacon Settings")
    Wait10s_ClickElement(data["menu_admin"]["add_Beacon"])
    Logging("- Add Beacon")

def TC_timesheet():
    Logging("------- Check menu crash - TimeCard -------")
    Wait10s_ClickElement(data["menu_timecard"]["button_timecard"])
    Logging("- Select time card")
    Wait10s_ClickElement(data["menu_timecard"]["button_timesheet"])
    Logging("- Select time sheet")

    Wait10s_ClickElement(data["menu_timecard"]["list"])
    Logging("- Tab List")
    time.sleep(5)

    # Check calendar crash
    Wait10s_ClickElement(data["menu_timecard"]["calendar_next"])
    Logging("- View next month")   
    time.sleep(5)
    Wait10s_ClickElement(data["menu_timecard"]["calendar_prev"])
    Logging("- View preview month")
    time.sleep(5)
    Wait10s_ClickElement(data["menu_timecard"]["calendar"])
    Wait10s_ClickElement(data["menu_timecard"]["select_date"])
    Logging("- Select date from calendar")
    time.sleep(5)

    list_sort_by = Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["sort"])
    Logging("- Sort by")
    time.sleep(2)
    list_week = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_timecard"]["list_tab"]["list_sort"])))
    if list_week.is_displayed():
        Logging("- Show list week")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)

    Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_2"])
    Logging("- 2nd Week")
    time.sleep(5)
        
    total_week_1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_timecard"]["list_tab"]["week_2_text"])))
    if total_week_1.text == 'TOTAL OF 2ND WEEK':
        Logging("=> TOTAL OF 2ND WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)

    list_sort_by.click
    Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_3"])
    Logging("- 3rd Week")
    time.sleep(5)

    total_week_2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_timecard"]["list_tab"]["week_3_text"])))
    if total_week_2.text == 'TOTAL OF 3RD WEEK':
        Logging("=> TOTAL OF 3RD WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)

    list_sort_by.click
    Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_4"])
    Logging("- 4th Week")
    time.sleep(5)
    total_week_3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_timecard"]["list_tab"]["week_4_text"])))
    if total_week_3.text == 'TOTAL OF 4TH WEEK':
        Logging("=> TOTAL OF 4TH WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)
        
    list_sort_by.click
    Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_5"])
    Logging("- 5th Week")
    time.sleep(5)
    total_week_4 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_timecard"]["list_tab"]["week_5_text"])))
    if total_week_4.text == 'TOTAL OF 5TH WEEK':
        Logging("=> TOTAL OF 5TH WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)
        
    Logging(" ")
    Logging("- Timesheet - Calendar -")
    Wait10s_ClickElement(data["menu_timecard"]["tab_calendar"])
    Logging("- Tab Calendar")
    time.sleep(5)

    # Check calendar crash
    Wait10s_ClickElement(data["menu_timecard"]["calendar_next"])
    Logging("- View next month")   
    time.sleep(5)
    Wait10s_ClickElement(data["menu_timecard"]["calendar_prev"])
    Logging("- View preview month")
    time.sleep(5)
    Wait10s_ClickElement(data["menu_timecard"]["calendar"])
    Wait10s_ClickElement(data["menu_timecard"]["select_date"])
    Logging("- Select date from calendar")
    time.sleep(5)

    
    driver.find_element_by_xpath(data["event"]["timecard"]).click()



Logging("Như Quỳnh")
execution()