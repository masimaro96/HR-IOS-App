from asyncio import events
from distutils.log import Log
import time, sys, unittest, random, json, openpyxl, platform
from xml.dom.minidom import Element
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
# from framework_sample import *

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

start_time = time.time()

desired_capabilities = {
    "xcodeOrgId": "9689HPSFXL",
    "xcodeSigningId": "iPhone Developer",
    "deviceName": "Hanbiro Iphone",
    "platformName": "IOS",
    "orientation": "PORTRAIT",
    "newCommandTimeout": 0,
    "automationName": "XCUITest",
    "derivedDataPath" : "/Users/hanbiro/Library/Developer/Xcode/DerivedData/WebDriverAgent-aghlrsejdreqngftgvcqwnjgrbou",
    "wdaConnectionTimeout": 500000,
    "udid": udid,
    "app": app_path,
    "bundleID": "com.hanbiro.RNGlobalHR",
    "wdaLaunchTimeout": 300000,
    "waitForQuiescene": False,
    "waitForIdleTimeout": 0,
    "autoAcceptAlerts": True
    #"showXcodeLog": True,
    #"showIOSLog": True
}

driver = webdriver.Remote(command_executor, desired_capabilities)

end_time = time.time()
duration = end_time - start_time
print("duration: %s" % duration)

n = random.randint(1,3000)

class Commands():
    def Wait10s_ClickElement(xpath):
        '''• Usage: Wait until the element visible and do the click
                return WebElement'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.click()

        return element

    def InputElement(xpath, value):
        '''• Usage: Send key value in input box
                return WebElement'''

        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value)

        return element
    
    def InputElement_2Values(xpath, value1, value2):
        '''• Usage: Send key with 2 values in input box
                return WebElement'''

        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value1)
        element.send_keys(value2)

        return element

    def Wait10s_InputElement(xpath, value):
        '''• Usage: Wait until the input box visible and send key value
                return WebElement'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value)

        return element

    def Wait10s_EnterElement(xpath, value):
        '''• Usage: Wait until the input box visible and send key value
            return WebElement'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value)
        element.send_keys(Keys.ENTER)

        return element

    def Wait10s_Clear_InputElement(xpath, value):
        '''• Usage: Wait until the input box visible and send key value
                return WebElement'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.clear()
        element.send_keys(value)

        return element

    def Wait10s_Clear_Click_InputElement(xpath, value):
        '''• Usage: Wait until the input box visible and send key value
            return WebElement'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.clear()
        element.click()
        element.send_keys(value)

        return element

class Waits():
    def WaitElementLoaded(time, xpath):
        '''• Usage: Wait until element VISIBLE in a selected time period'''
        
        WebDriverWait(driver, time).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element

    def Wait10s_ElementLoaded(xpath):
        '''• Usage: Wait 10s until element VISIBLE'''
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element

    def WaitElementInvisibility(time, xpath):
        '''• Usage: Wait until element INVISIBLE in a selected time period'''
        
        WebDriverWait(driver, time).until(EC.invisibility_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element

    def Wait10s_ElementInvisibility(xpath):
        '''• Usage: Wait 10s until element INVISIBLE'''
        
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element
    
    def WaitUntilPageIsLoaded(page_xpath):
        if bool(page_xpath) == True:
            # wait until page's element is present
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, page_xpath)))

        # check if the loading icon is not present at the page -> page is completely loaded
        try:
            WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='loading-dialog hide']")))
        except WebDriverException:
            pass

def execution():
    time.sleep(10)
    Logging("------- Login to app -------")
    # Input information for log-in
    driver.find_element_by_ios_class_chain(data["domain"]).send_keys(data["domain_name"])
    Logging("- Input Domain")
    driver.find_element_by_ios_class_chain(data["id_login"]).send_keys(data["id_login_name_1"])
    Logging("- Input ID")
    driver.find_element_by_ios_class_chain(data["pass_login"]).send_keys(data["password"])
    Logging("- Input Password")
    Commands.Wait10s_ClickElement(data["login_button"])
    Logging("=> Click Log In button")
    #driver.implicitly_wait(1000)

    #add_event()
    #Commands.Wait10s_ClickElement("Timecard")
    #element = Commands.Wait10s_ClickElement("//*[contains(., 'Timesheet')]")
    #Waits.Wait10s_ElementLoaded("//*[contains(., 'Work Policy')]")
    time.sleep(20)
    clock_in()
    admin_settings_GPS()
    admin_settings_wifi()

def clock_in():
    reason_OT = data["OT"]["reason_text"]
    reason_late = data["clock_in"]["reason_late_text"]
    try:
        Logging("--- Clock in with GPS ---")
        try:
            OT = Waits.Wait10s_ElementLoaded(data["OT"]["status_OT"])   
            if OT.text == 'Night shift':
                Logging("=> Work night shift")
                Commands.Wait10s_ClickElement(data["OT"]["confirm_OT"])
                Commands.Wait10s_ClickElement(data["OT"]["apply_OT"])
                Logging("=> Confirm OT")

                Commands.Wait10s_InputElement(data["OT"]["reason_OT"], reason_OT)
                Logging("=> Input reason OT")

                Commands.Wait10s_ClickElement(data["OT"]["confirm_apply_OT"])
                Commands.Wait10s_ClickElement(data["OT"]["button_close"])
                Logging("=> Apply OT success")
            else:
                Logging("=> Apply OT not display")
        except WebDriverException:
            Logging("=> Apply OT not display")
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["clock_in"]["button_clock_in"])
        Logging("=> Click clock in")
        time.sleep(5)
        status_clock_in= Waits.Wait10s_ElementLoaded(data["clock_in"]["status_late"])
        if status_clock_in.text == 'Tardiness':
            Logging("=> Clock in late")
            Commands.Wait10s_InputElement(data["clock_in"]["reason_late"], reason_late)
            Logging("=> Input reason late")
            time.sleep(5)
            Commands.Wait10s_ClickElement(data["clock_in"]["status_late"])
            Commands.Wait10s_ClickElement(data["clock_in"]["button_save"])
            Logging("=> Save")
        else:
            Logging("=> Clock in on time")
            Commands.Wait10s_ClickElement(data["clock_in"]["button_close"])

    except WebDriverException:
        Logging("=> Crash app")

def break_time():
    try:
        Logging("--- Break time ---")
        Commands.Wait10s_ClickElement(data["break_time"]["button_break_time"])
        Logging("=> Start break time")
        time.sleep(30)
        Commands.Wait10s_ClickElement(data["break_time"]["end_break_time"])
        Logging("=> End Breake time")
        time.sleep(10)
    except WebDriverException:
        Logging("=> User have clock out")

def clock_out():
    reason_key = data["clock_out"]["reason_clock_out_text"]
    try:
        Logging("--- Clock out ---")
        Commands.Wait10s_ClickElement(data["clock_out"]["button_clock_out"])
        Logging("=> Click clock out")
        time.sleep(20)
        status = Waits.Wait10s_ElementLoaded(data["clock_out"]["status_clock_out"])
        if status.text == 'Leave Early':
            Logging("=> Clock out early")
            Commands.Wait10s_InputElement(data["clock_out"]["reason_clock_out"], reason_key)
            Logging("=> Input reason")
            time.sleep(10)
            Commands.Wait10s_ClickElement(data["clock_out"]["confirm_clock_out"])
            Commands.Wait10s_ClickElement(data["clock_out"]["clock_out_success"])
            Logging("=> Clock out success")
        else:
            Logging("=> Clock out on time")
            Commands.Wait10s_ClickElement(data["clock_out"]["button_close"])
    except WebDriverException:
        Logging("=> Crash app")

def view_calendar():
    Commands.Wait10s_ClickElement(data["calender_home"]["date"])
    Logging("- Select calender")
    Commands.Wait10s_ClickElement(data["calender_home"]["date_to_view"])
    Commands. Wait10s_ClickElement(data["calender_home"]["select"])
    Logging("- Select date")

    popup = Waits.WaitElementLoaded(20, data["calender_home"]["popup_warning"])
    if popup == "Warning":
        Logging("- Date select not worked day -> Select again")
    else:
        Logging("- Date correct")

    Commands.Wait10s_ClickElement(data["calender_home"]["preview_date"])
    Logging("- View preview date")
    Commands.Wait10s_ClickElement(data["calender_home"]["next_date"])
    Logging("- View next date")

def view_noti():
    Logging("--- View Notification ---")
    Commands.Wait10s_ClickElement(data["menu_settings"]["button_notification"])
    Logging("- View content notification")
    Commands.Wait10s_ClickElement(data["menu_settings"]["button_back"])
    Logging("-> Back to menu")

def settings():
    try:
        Logging("--- Settings - Change language ---")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        Commands.Wait10s_ClickElement("//*[@text='한국어']")
        Commands.Wait10s_ClickElement("//*[@text='닫기']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        
        korean_text = Waits.Wait10s_ElementLoaded(data["settings"]["language_text"])
        if korean_text.text == '한국어':
            Logging("=> Change to language '한국어' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='언어 설정']")
        
        Commands.Wait10s_ClickElement("//*[@text='Tiếng Việt']")
        Commands.Wait10s_ClickElement("//*[@text='Đóng']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        VN_text = Waits.Wait10s_ElementLoaded(data["settings"]["language_text"])
        if VN_text.text == 'Tiếng Việt':
            Logging("=> Change to language 'Tiếng Việt' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='Thay đổi ngôn ngữ']")

        Commands.Wait10s_ClickElement("//*[@text='日本語']")
        Commands.Wait10s_ClickElement("//*[@text='閉じる']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        JP = Waits.Wait10s_ElementLoaded(data["settings"]["language_text"])
        if JP.text == '日本語':
            Logging("=> Change to language '日本語' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='言語']")

        Commands.Wait10s_ClickElement("//*[@text='简体中文']")
        Commands.Wait10s_ClickElement("//*[@text='關閉']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        TQ = Waits.Wait10s_ElementLoaded(data["settings"]["language_text"])
        if TQ.text == '简体中文':
            Logging("=> Change to language '简体中文' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='語言']")

        Commands.Wait10s_ClickElement("//*[@text='Indonesian']")
        Commands.Wait10s_ClickElement("//*[@text='Tutup']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        indo = Waits.Wait10s_ElementLoaded(data["settings"]["language_text"])
        if indo.text == 'Indonesian':
            Logging("=> Change to language 'Indonesian' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='Ganti BAHASA']")

        Commands.Wait10s_ClickElement("//*[@text='English']")
        Commands.Wait10s_ClickElement("//*[@text='Close']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        EN = Waits.Wait10s_ElementLoaded(data["settings"]["language_text"])
        if EN.text == 'English':
            Logging("=> Change to language 'English' success")
        else:
            Logging("=> Fail")
        
    except WebDriverException:
        Logging("=> Crash app")

def add_event():
    title = data["event"]["title_text"]
    location = data["event"]["location_text"]
    memo = data["event"]["memo_text"]
    try:
        Logging(" ")
        Logging("------- Add event -------")
        Commands.Wait10s_ClickElement(data["menu_timecard"]["button_timecard"])
        Logging("- Select time card")
        Commands.Wait10s_ClickElement(data["menu_timecard"]["button_timesheet"])
        Logging("- Select time sheet")
        Commands.Wait10s_ClickElement(data["event"]["add_event"])
        Logging("- Select add")
        Commands.Wait10s_InputElement("//*[contains(@text,'Please input data.')]", title)
        Logging("- Input title")
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["event"]["choose_event"])
        Commands.Wait10s_ClickElement(data["event"]["type_event"])
        Logging("- Choose event")
        Commands.Wait10s_ClickElement(data["event"]["select_color"])
        Logging("- Select color")
        Commands.Wait10s_InputElement(data["event"]["place"], location)
        Logging("- Input location")
        Commands.Wait10s_InputElement(data["event"]["memo"], memo)
        Logging("- Input memo")
        Commands.Wait10s_ClickElement(data["event"]["button_save"])
    except:
        Logging("- Can't create event")

    Logging("** Check event use approval type")
    try:
        approval_type = Waits.Wait10s_ElementLoaded(data["event"]["popup"])
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

def admin_settings_GPS():
    Commands.Wait10s_ClickElement(data["menu_admin"]["button_admin"])
    Logging("- Admin settings")

    Commands.Wait10s_ClickElement(data["menu_admin"]["GPS_setting"])
    Logging("- GPS Settings")
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["add_GPS"])
    time.sleep(10)
    driver.find_element_by_ios_class_chain(data["menu_admin"]["popup"]).click()
    Logging("- Add GPS")
    Commands.Wait10s_ClickElement(data["menu_admin"]["search_gps"])
    Logging("- Input GPS")
    time.sleep(10)
    Commands.InputElement(data["menu_admin"]["search_gps"], "Nguyen")
    Logging("- Enter value")
    Commands.Wait10s_ClickElement(data["menu_admin"]["done"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["gps"])
    Logging("- Select GPS")
    Commands.Wait10s_ClickElement(data["menu_admin"]["workplace"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["select_workplace"])
    Logging("- Select workplace")
    time.sleep(10)
    Commands.Wait10s_ClickElement(data["menu_admin"]["save_GPS"])
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["close_popup"])
    Logging("- Save GPS")
    time.sleep(5)
    driver.swipe(start_x=286, start_y=140, end_x=100, end_y=140, duration=800)
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["delete_gps"])
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["button_yes"])
    Logging("- Delete GPS")
    Commands.Wait10s_ClickElement(data["menu_admin"]["close_popup"])

def admin_settings_wifi():
    Commands.Wait10s_ClickElement(data["menu_admin"]["back_button"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["Wifi_setting"])
    Logging("- Wifi Settings")
    Commands.Wait10s_ClickElement(data["menu_admin"]["add_GPS"])
    Logging("- Add Wifi")
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["add_wifi"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["next_button"])
    Logging("- Choose Wifi")
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["workplace_wifi"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["select_workplace_wifi"])
    Logging("- Choose Workplace")
    Commands.Wait10s_ClickElement(data["menu_admin"]["next"])
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["close_popup"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["complete"])
    Logging("- Save wifi")
    time.sleep(5)
    driver.swipe(start_x=286, start_y=140, end_x=100, end_y=140, duration=800)
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["delete_gps"])
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_admin"]["button_yes"])
    Logging("- Delete wifi")
    Commands.Wait10s_ClickElement(data["menu_admin"]["close_popup"])

def admin_settings_beacon():
    Commands.Wait10s_ClickElement(data["menu_admin"]["back_button"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["Beacon_setting"])
    Logging("- Beacon Settings")
    Commands.Wait10s_ClickElement(data["menu_admin"]["add_Beacon"])
    Logging("- Add Beacon")

def select_date_month():
    Commands.Wait10s_ClickElement(data["next"])
    Logging("- View next date-month")
    Commands.Wait10s_ClickElement(data["prev"])
    Logging("- View pre date-month")
    Commands.Wait10s_ClickElement(data["calendar_select"])
    time.sleep(5)
    Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=19]")
    time.sleep(5)
    dateselect = Waits.Wait10s_ElementLoaded(data["TimeCard"]["weekly_status"]["date_calendar"])
    dateselect_text = dateselect.text
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["select_date"])
    Logging("- Select calendar")

    return dateselect_text

def input_user_name():
    #driver.find_element_by_ios_class_chain(data["domain_input"]) "//*[@text='Please insert keyword to search']")

    ''' Send key "quynh" from keyboard mobile '''
    driver.is_keyboard_shown()
    Commands.Wait10s_ClickElement("//XCUIElementTypeButton[@name='q']")
    Commands.Wait10s_ClickElement("//XCUIElementTypeButton[@name='u']")
    Commands.Wait10s_ClickElement("//XCUIElementTypeButton[@name='y']")
    Commands.Wait10s_ClickElement("//XCUIElementTypeButton[@name='n']")
    Commands.Wait10s_ClickElement("//XCUIElementTypeButton[@name='h']")
    Commands.Wait10s_ClickElement("//XCUIElementTypeButton[@name='Done']")
    Logging("- Search user")

def TC_timesheet():
    Logging("------- Check menu crash - TimeCard -------")
    Commands.Wait10s_ClickElement(data["menu_timecard"]["button_timecard"])
    Logging("- Select time card")
    Commands.Wait10s_ClickElement(data["menu_timecard"]["button_timesheet"])
    Logging("- Select time sheet")

    Commands.Wait10s_ClickElement(data["menu_timecard"]["list"])
    Logging("- Tab List")
    time.sleep(5)

    # Check calendar crash
    Commands.Wait10s_ClickElement(data["menu_timecard"]["calendar_next"])
    Logging("- View next month")   
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_timecard"]["calendar_prev"])
    Logging("- View preview month")
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_timecard"]["calendar"])
    Commands.Wait10s_ClickElement(data["menu_timecard"]["select_date"])
    Logging("- Select date from calendar")
    time.sleep(5)

    list_sort_by = Commands.Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["sort"])
    Logging("- Sort by")
    time.sleep(2)
    list_week = Commands.Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["list_sort"])
    if list_week.is_displayed():
        Logging("- Show list week")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)

    Commands.Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_2"])
    Logging("- 2nd Week")
    time.sleep(5)
        
    total_week_1 = Waits.Wait10s_ElementLoaded(data["menu_timecard"]["list_tab"]["week_2_text"])
    if total_week_1.text == 'TOTAL OF 2ND WEEK':
        Logging("=> TOTAL OF 2ND WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)

    list_sort_by.click
    Commands.Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_3"])
    Logging("- 3rd Week")
    time.sleep(5)

    total_week_2 = Waits.Wait10s_ElementLoaded(data["menu_timecard"]["list_tab"]["week_3_text"])
    if total_week_2.text == 'TOTAL OF 3RD WEEK':
        Logging("=> TOTAL OF 3RD WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)

    list_sort_by.click
    Commands.Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_4"])
    Logging("- 4th Week")
    time.sleep(5)
    total_week_3 = Waits.Wait10s_ElementLoaded(data["menu_timecard"]["list_tab"]["week_4_text"])
    if total_week_3.text == 'TOTAL OF 4TH WEEK':
        Logging("=> TOTAL OF 4TH WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)
        
    list_sort_by.click
    Commands.Wait10s_ClickElement(data["menu_timecard"]["list_tab"]["week_5"])
    Logging("- 5th Week")
    time.sleep(5)
    total_week_4 = Waits.Wait10s_ElementLoaded(data["menu_timecard"]["list_tab"]["week_5_text"])
    if total_week_4.text == 'TOTAL OF 5TH WEEK':
        Logging("=> TOTAL OF 5TH WEEK")
    else:
        Logging("=> Crash app")
        exit(0)
    time.sleep(5)
        
    Logging(" ")
    Logging("- Timesheet - Calendar -")
    Commands.Wait10s_ClickElement(data["menu_timecard"]["tab_calendar"])
    Logging("- Tab Calendar")
    time.sleep(5)

    # Check calendar crash
    Commands.Wait10s_ClickElement(data["menu_timecard"]["calendar_next"])
    Logging("- View next month")   
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_timecard"]["calendar_prev"])
    Logging("- View preview month")
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["menu_timecard"]["calendar"])
    Commands.Wait10s_ClickElement(data["menu_timecard"]["select_date"])
    Logging("- Select date from calendar")
    time.sleep(5)

def TC_report():
    try:
        driver.find_element_by_xpath(data["event"]["timecard"]).click()
        Commands.Wait10s_ClickElement(data["TimeCard"]["report_monthly"]["MT_report"])
        Logging("- Schedule Working")
        time.sleep(10)

        schedule = Waits.Wait10s_ElementLoaded(data["TimeCard"]["report_monthly"]["schedule_working"])
        if schedule.text == 'Scheduled working day':
            count_day = Waits.Wait10s_ElementLoaded(data["TimeCard"]["report_monthly"]["count_schedule_working"])
            Logging("- Scheduled working day:", count_day.text)
        else:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["report_monthly"]["events"])
        Logging("- Events")
        time.sleep(5)
        clock_in = Waits.Wait10s_ElementLoaded(data["TimeCard"]["report_monthly"]["clockin"])
        if clock_in.text == 'Clock-In':
            count_clock_in = Waits.Wait10s_ElementLoaded(data["TimeCard"]["report_monthly"]["count_clockin"])
            Logging("- Events - Clock in:", count_clock_in.text)
        else:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["report_monthly"]["working_status"])
        Logging("- Working status")
        time.sleep(5)
        working_time = Waits.Wait10s_ElementLoaded(data["TimeCard"]["report_monthly"]["workingtime"])
        if working_time.text == 'Working time':
            count_working_time = Waits.Wait10s_ElementLoaded(data["TimeCard"]["report_monthly"]["count_workingtime"])
            Logging("- Working status - Working time:", count_working_time.text)
        else:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)

        Logging("** Check report - Weekly")
        Commands.Wait10s_ClickElement(data["TimeCard"]["report_weekly"]["weekly"])
        Logging("- View week tab")
        time.sleep(10)

        Commands.Wait10s_ClickElement("//*[@text='Device']")
        Logging("- View tab device")
        time.sleep(5)

        driver.swipe(start_x=1174, start_y=730, end_x=400, end_y=730, duration=800)
        time.sleep(5)

        Commands.Wait10s_ClickElement("//*[@text='Average working hour per week']")
        Logging("- View tab Avg_Working")
        time.sleep(5)
        driver.swipe(start_x=1174, start_y=730, end_x=400, end_y=730, duration=800)

        Commands.Wait10s_ClickElement("//*[@text='Working hours per day of the week']")
        Logging("- View tab working hour")
        time.sleep(5)
        
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["next"])
        Logging("- View next date")   
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["prev"])
        Logging("- View preview date")
        time.sleep(5)
        
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["calendar_select"])
        Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=9]")
        Logging("- Select date from calendar")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_list"]["select_date"])
        time.sleep(5)

        Logging("** Check report - List")
        Commands.Wait10s_ClickElement(data["TimeCard"]["report_list"]["list"])
        Logging("- View list tab")
        time.sleep(10)
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["next"])
        Logging("- View next date")   
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["prev"])
        Logging("- View preview date")
        time.sleep(5)
        
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["calendar_select"])
        Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=9]")
        Logging("- Select date from calendar")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_list"]["select_date"])
        time.sleep(5)
    except:
        Logging("-> Crash app")

def CP_daily_status():
    try:
        Logging(" ")
        Logging("** Check Company timecard - Daily status")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["next"])
        Logging("- View next date")   
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["prev"])
        Logging("- View preview date")
        time.sleep(5)
        
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_calendar"]["calendar_select"])
        Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=9]")
        Logging("- Select date from calendar")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timesheet_list"]["select_date"])
        time.sleep(5)
    except:
        Logging("-> Crash app")

def CP_weekly_status():
    try:
        Logging(" ")
        Logging("** Check Company timecard - Weekly Status")
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["weekly_status_CT"])
        time.sleep(5)
        try:
            title = Waits.Wait10s_ElementLoaded(data["TimeCard"]["weekly_status"]["week_title"])
            if title.text == 'Weekly Status':
                Logging("- Show content")

        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        try:
            Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["sort_week"])
            Logging("- Select sort week")
            sortweek = Waits.Wait10s_ElementLoaded(data["TimeCard"]["weekly_status"]["list"])
            if sortweek.is_displayed:
                Logging("- List week display")
            else:
                Logging("- List not display")
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        Logging("- Select week")
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["week_2"])
        Logging("- 2nd Week")
        time.sleep(5)
        try:
            user_view = Waits.Wait10s_ElementLoaded(data["TimeCard"]["weekly_status"]["user"])
            if user_view.is_displayed():
                Logging("- Show content")
            else:
                Logging("- No data")
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["sort_week"])
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["week_3"])
        Logging("- 3rd Week")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["sort_week"])
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["week_4"])
        Logging("- 4th Week")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["sort_week"])
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["week_5"])
        Logging("- 5th Week")
        time.sleep(5)

        '''Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["sort_week"])
        time.sleep(5)
        driver.swipe(start_x=291, start_y=1126, end_x=291, end_y=900, duration=800)
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["week_6"])
        Logging("- 6th Week")
        time.sleep(5)'''

        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["sort_week"])
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["week_1"])
        Logging("- 1st Week")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["search"])
        input_user_name()

        Commands.Wait10s_ClickElement("//*[contains(@text,'quynh1')]")
        time.sleep(5)

        ''' Next - Prev - Select date '''
        Commands.Wait10s_ClickElement(data["next"])
        Logging("- View next date-month")
        Commands.Wait10s_ClickElement(data["prev"])
        Logging("- View pre date-month")
        Commands.Wait10s_ClickElement(data["calendar_select"])
        time.sleep(5)
        Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=19]")
        time.sleep(5)
        dateselect = Waits.Wait10s_ElementLoaded(data["TimeCard"]["weekly_status"]["date_calendar"])
        dateselect_text = dateselect.text
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["select_date"])
        Logging("- Select calendar")

        time.sleep(10)

        try:
            user_view = Waits.Wait10s_ElementLoaded(data["TimeCard"]["weekly_status"]["user"])
            if user_view.is_displayed():
                Logging("- Show content")
            else:
                Logging("- No data")
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        
        Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["view_detail"])
        Logging("- View detail")

        date_of_calendar = Waits.Wait10s_ElementLoaded(data["TimeCard"]["weekly_status"]["date"])
        date = date_of_calendar.text
        x = date.split(" ")[2]
        a = x.split(",")[0]

        if a == dateselect_text:
            Logging("- Show right date")
        else:
            Logging("- Crash")
            exit(0)
    except:
        Logging("-> Crash app")

    Commands.Wait10s_ClickElement(data["TimeCard"]["weekly_status"]["back"])

def CP_timeline():
    try:
        Logging(" ")
        Logging("** Check Company timecard - Time Line")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["timeline_CT"])
        time.sleep(5)
        try:
            title = Waits.Wait10s_ElementLoaded(data["TimeCard"]["timeline"]["timeline_title"])
            if title.text == 'Time Line':
                Logging("- Show content")

        except WebDriverException:
            Logging("=> Crash app")
            exit(0)

        ''' Search user '''
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["search"])
        input_user_name()

        Commands.Wait10s_ClickElement("//*[contains(@text,'quynh1')]")
        Logging("- Select user")
        time.sleep(5)

        ''' Next - Prev - Select date '''
        select_date_month()
        time.sleep(10)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        Logging("- Sort time line")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["clockin"])
        Logging("- Clock in")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        Logging("- Sort time line")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["tardin"])
        Logging("- Tardines")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        Logging("- Sort time line")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["clockout"])
        Logging("- Clock out")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        Logging("- Sort time line")
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["auto_clockout"])
        Logging("- Automatically Clock-out")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        Logging("- Sort time line")
        time.sleep(5)
        driver.swipe(start_x=291, start_y=1315, end_x=291, end_y=691, duration=800)
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["outside"])
        Logging("- Outside")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        time.sleep(5)
        driver.swipe(start_x=291, start_y=1315, end_x=291, end_y=691, duration=800)
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["meeting"])
        Logging("- Meeting")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        time.sleep(5)
        driver.swipe(start_x=291, start_y=1315, end_x=291, end_y=691, duration=800)
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["edu"])
        Logging("- Education")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        time.sleep(5)
        driver.swipe(start_x=291, start_y=1315, end_x=291, end_y=691, duration=800)
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["business"])
        Logging("- Business Trip")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        time.sleep(5)
        driver.swipe(start_x=291, start_y=1315, end_x=291, end_y=691, duration=800)
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["working_remote"])
        Logging("- Working remote")
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["sort_timeline"])
        time.sleep(5)
        driver.swipe(start_x=291, start_y=1315, end_x=291, end_y=691, duration=800)
        driver.swipe(start_x=291, start_y=1315, end_x=291, end_y=691, duration=800)
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["timeline"]["working"])
        Logging("- Working")
        time.sleep(5)
    except:
        Logging("-> Crash app")

    Commands.Wait10s_ClickElement(data["event"]["timecard"]) 

def CP_report():
    try:
        Logging(" ")
        Logging("** Check Company timecard - Report")
        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["report_CT"])
        Logging("- View by work")
        time.sleep(5)
        try:
            title = Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["report_title"])
            if title.text == 'Report':
                Logging("- Show content")
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)

        ''' Next - Prev - Select date '''
        select_date_month()
        time.sleep(10)

        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["search"])
        Logging("- Search user")
        input_user_name()
        time.sleep(3)
        Commands.Wait10s_ClickElement("//*[contains(@text,'quynh1')]")
        Logging("- Select user")
        time.sleep(5)
        
        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["detail"])
        Logging("- View detail")
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["back"])
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["view_by_event"])
        Logging("- View by event")
        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["detail"])
        Logging("- View detail")
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["back"])
        time.sleep(5)
    except:
        Logging("-> Crash app")

    Commands.Wait10s_ClickElement(data["event"]["timecard"])

    try:
        Logging(" ")
        Logging("** Check Company timecard - Approval")
        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["approval_CT"])
        time.sleep(5)
        try:
            title = Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["approval_title"])
            if title.text == 'Approval':
                Logging("- Show content")
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        time.sleep(10)

        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["filter_type"])
        Logging("- Filter type")
        time.sleep(5)
        try:
            filter = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["detail"])
            if filter.text == 'Detail':
                Logging("-")
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)

        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["type"])
        Logging("- Select type")
        time.sleep(5)
        try:
            type_detail = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["event"])
            if type_detail.is_displayed():
                type_detail.click()
                time.sleep(5)
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        
        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["status"])
        Logging("- Select status")
        time.sleep(5)
        try:
            status = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["pending"])
            if status.is_displayed():
                status.click()
                time.sleep(5)
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)

        '''Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["peroid"])
        Logging("- Select peroid")
        time.sleep(5)
        try:
            peroid = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["TimeCard"]["approval"]["today"])
            if peroid.is_displayed():
                peroid.click()
                time.sleep(5)
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)'''

        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["back"])
        time.sleep(5)
        try:
            title = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["approval_title"])
            if title.text == 'Approval':
                Logging("- Show content")
        except WebDriverException:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["report"]["search"])
        
        input_user_name()

        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["user"])
        #WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//android.widget.TextView[@text='quynh1']")
        Logging("- Select user")
        time.sleep(5)

        ''' Approve '''
        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["select"])
        try:
            approve_line = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["approval_line"])
            if approve_line.is_displayed():
                Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["approve"])
                time.sleep(5)
                Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["close"])
                Logging("- Approve request")
                time.sleep(5)
                
                Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["filter_type"])
                Logging("- Filter type")
                time.sleep(5)
                Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["status"])
                Logging("- Select status")
                time.sleep(5)
                Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["approved"])
                time.sleep(5)

                Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["back"])
                time.sleep(5)

                status_approve = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["status_text"])
                if status_approve.text == 'Approved':
                    Logging("- Approve success")
                else:
                    Logging("- Fail")
            else:
                Logging("=> Approve don't have approve permission")
        except WebDriverException:
            Logging("=> Approve don't have approve permission")

        ''' Reject '''
        add_event()
        Commands.Wait10s_ClickElement(data["event"]["timecard"])
        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["approval_CT"])
        time.sleep(5)
        Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["select"])
        try:
            approve_line = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["approval_line"])
            if approve_line.is_displayed():
                Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["reject_bt"])
                time.sleep(5)
                Commands.Wait10s_ClickElement(data["TimeCard"]["approval"]["close"])
                time.sleep(5)
                status_reject = Waits.Wait10s_ElementLoaded(data["TimeCard"]["approval"]["status_text"])
                if status_reject.text == 'Rejected':
                    Logging("- Rejected success")
                else:
                    Logging("- Fail")
            else:
                Logging("=> Approve don't have approve permission")
        except WebDriverException:
            Logging("=> Approve don't have approve permission")
    except:
        Logging("-> Crash app")

def attachfile():
    try:
        Commands.Wait10s_ClickElement(data["attach_file"]["add_file"])
        Logging("- Select attach file")
        Commands.Wait10s_ClickElement(data["attach_file"]["choose_photo"])
        Logging("- Choose photo")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, data["attach_file"]["choose_gallery"]))).click()
        Logging("- Choose gallery")
        Commands.Wait10s_ClickElement(data["attach_file"]["select_photo"])
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, data["attach_file"]["select"]))).click()
        Logging("- Select photo")
        time.sleep(5)
    except:
        Logging("- Can't attach file")

def request_vacation():
    Commands.Wait10s_ClickElement(data["menu_timecard"]["button_timecard"])
    Commands.Wait10s_ClickElement(data["vacation"]["button_vacation"])
    
    title_request = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["request_vacation_text"])
    if title_request.text == 'Request vacation':
        Logging("=> Request vacation")
    else:
        Logging("=> Crash app")
        exit(0)   
    try:
        vacation_type = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["AM"])
        if vacation_type.is_displayed():
            vacation_type.click()
            Logging("- Select vacation type")

            Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["calendar"])
            Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]")
            Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]")
            time.sleep(2)
            Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["select_calendar"])
            time.sleep(2)

            ''' Crash app when select date '''
            try:
                title_request = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["request_vacation_text"])
                if title_request.text == 'Request vacation':
                    Logging("- Select date vacation")
                else:
                    Logging("=> Crash app")
                    exit(0)  
            except WebDriverException: 
                Logging("=> Crash app")
                exit(0)

            ''' Get data of vacation request '''
            vacation_date = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["request_date_text"])
            vacation_text = vacation_date.text
            date_text = vacation_text.split(" ")[0]
            type_vacation = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["AM"])
            type_text = type_vacation.text
            vacation_date_type = date_text + "(" + type_text + ")"            
    except:
        Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["calendar"])
        Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]")
        Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]")
        time.sleep(2)
        Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["select_calendar"]) 
        time.sleep(2)

        ''' Get data of vacation request '''
        vacation_date = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["request_date_text"])
        vacation_text = vacation_date.text
        # date_text = vacation_text.split(" ")[0]
        # type_vacation = Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["AM"])))
        # type_text = type_vacation.text
        # vacation_date_type = vacation_text + "(" + type_text + ")"
        
    try:
        attachfile()
    except:
        pass

    driver.swipe(start_x=650, start_y=1844, end_x=650, end_y=355, duration=800)
    time.sleep(5)
    ''' Select CC '''
    CC = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["CC"])
    if CC.is_displayed():
        Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["CC"])
        input_user_name()
        Commands.Wait10s_ClickElement("//*[contains(@text,'quynh1')]")
        Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["save_cc"])
        Logging("- Select CC")
    else:
        driver.swipe(start_x=650, start_y=1662, end_x=650, end_y=355, duration=800)
        reason = data["vacation"]["my_vacation"]["input_test"]
        Commands.Wait10s_InputElement("//*[@text='Please enter your reason']", reason)
        Logging("- Input reason")
    
    Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["button_request"])
    
    '''- Check day request
      + If vacation day is saturday => fail, check again
      + If memo is empty => fail, check again'''
    try:
        fail = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["request_fail"])
        if fail.text == 'request vacation failure':
            Logging("--- Request vacation failure - vacation day is saturday---")
            Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["close_fail"])
            time.sleep(2)
            driver.swipe(start_x=650, start_y=355, end_x=650, end_y=2275, duration=800)
            Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["calendar"])
            time.sleep(2)
            Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]")
            Logging("=> Select start date")
            
            Commands.Wait10s_ClickElement("//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]")
            Logging("=> Select end date")
            Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["select_calendar"])
            Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["button_request"])
            Logging("=> Send request vacation")
        else:
            Logging("=> Request success")
    except WebDriverException:
        Logging("=> Request success") 

    Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["button_close"])

def log_in():
    ''' Log in '''
    id_user2 = data["id_name_2"]
    pass_user = data["pass_input"]
    Commands.Wait10s_Clear_InputElement(data["vacation"]["vacation_approve"]["user_name"], id_user2)
    Logging("- Input ID")
    #driver.find_element_by_ios_class_chain(data["domain_input"]) "//*[@text='Password']")))
    Commands.Wait10s_InputElement(data["password"], pass_user)
    Logging("- Input Password")
    Commands.Wait10s_ClickElement("//*[contains(@text,'Login')]")
    Commands.Wait10s_ClickElement(data["button_login"])
    Logging("=> Click Log In button")
    driver.implicitly_wait(50)

def approve_request():

    request_vacation()

    Logging(" ")
    ''' Log out '''
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["setting_button"])
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["logout"])
    Logging("=> Change to user 2 to approve request")
    log_in()

    ''' Check request vacation of user 1 '''
    time.sleep(5)
    Logging("- Check request vacation")
    Commands.Wait10s_ClickElement(data["vacation"]["button_vacation"])
    Commands.Wait10s_ClickElement(data["vacation"]["manage_processing"]["vacation_approve"])
    time.sleep(3)
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["search"])
    input_user_name()
    Logging("- Search user")

    Commands.Wait10s_ClickElement("//*[contains(@text,'quynh1')]")
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["select_user"])
    Logging("- Select user")
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["approve"])
    Logging("- Approve request")
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["accept_approve"])
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["close_popup"])
    Logging("=> Approve success")

    text_approve = Waits.Wait10s_ElementLoaded(data["vacation"]["vacation_approve"]["approve_text"])
    if text_approve.text == 'Approved':
        Logging("=> Request have approve success")
    else:
        Logging("=> Approve fail")

def cancel_request():
    ''' User cancel request '''
    Logging(" ")
    ''' Log out '''
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["setting_button"])
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["logout"])
    Logging("=> Change to user 1 - check request have been approve - cancel request")

    ''' Log in '''
    log_in()

    ''' Check request vacation of user 1 '''
    time.sleep(5)
    Logging("- Check request vacation")
    Commands.Wait10s_ClickElement(data["vacation"]["button_vacation"])
    Commands.Wait10s_ClickElement(data["vacation"]["manage_processing"]["vacation_approve"])
    time.sleep(3)
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["search"])
    input_user_name()

    Commands.Wait10s_ClickElement("//*[contains(@text,'quynh1')]")

    text_approve = Waits.Wait10s_ElementLoaded(data["vacation"]["vacation_approve"]["approve_text"])
    if text_approve.text == 'Approved':
        Logging("=> Request have approve success")
    else:
        Logging("=> Approve fail")

    Commands.Wait10s_ClickElement(data["vacation"]["button_vacation"])
    time.sleep(5)
    Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["vacation_status"]["vacationstatus"])
    time.sleep(10)
    Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["vacation_status"]["cancel_request"])
    Commands.Wait10s_ClickElement(data["vacation"]["my_vacation"]["vacation_status"]["button_ok"])
    Logging("- Cancel request")
    time.sleep(10)
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["close_popup"])
    Logging("=> Approve cancel request success")
    time.sleep(5)

    text_cancel = Waits.Wait10s_ElementLoaded(data["vacation"]["my_vacation"]["vacation_status"]["text_request"])
    if text_cancel.text == 'User cancel':
        Logging("=> Send cancel request success")
    else:
        Logging("=> Approve Arbitrary decision")

def apporve_cancel_request():
    Logging(" ")
    ''' Log out '''
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["setting_button"])
    Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["logout"])
    Logging("=> Change to user 2 - approve cancel request")
    log_in()

    try:
        Logging("- Check request vacation")
        Commands.Wait10s_ClickElement(data["vacation"]["button_vacation"])
        Commands.Wait10s_ClickElement(data["vacation"]["manage_processing"]["vacation_approve"])
        time.sleep(3)
        Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["cancel_request"])
        Logging("- Click cancel request")
        Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["status"])
        Commands.Wait10s_ClickElement("//*[@text='Request']")
        Logging("- Select status request")
        Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["request"])
        Logging("- Select request")
        Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["approve_cancel"])
        Logging("- APPROVE CANCELLATION")
        Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["accept_approve_cancel"])
        Logging("=> Cancel request")
        Commands.Wait10s_ClickElement(data["vacation"]["vacation_approve"]["close_popup"])
    except:
        Logging("=> Cancel request fail")