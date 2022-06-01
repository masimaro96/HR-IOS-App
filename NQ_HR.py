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
    reason_OT = data["OT"]["reason_text"]
    reason_late = data["clock_in"]["reason_late_text"]
    try:
        Logging("--- Clock in with GPS ---")
        try:
            OT = Waits.Wait10s_ElementInvisibility(data["OT"]["status_OT"])   
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

        Commands.Wait10s_ClickElement(data["clock_in"]["button_clock_in"])
        Logging("=> Click clock in")

        status_clock_in= Waits.Wait10s_ElementInvisibility(data["clock_in"]["status_late"])
        if status_clock_in.text == 'Tardiness':
            Logging("=> Clock in late")
            Commands.Wait10s_InputElement(data["clock_in"]["reason_late"], reason_late)
            Logging("=> Input reason late")
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

        status = Waits.Wait10s_ElementInvisibility(data["clock_out"]["status_clock_out"])
        if status.text == 'Leave Early':
            Logging("=> Clock out early")
            Commands.Wait10s_InputElement(data["clock_out"]["reason_clock_out"], reason_key)
            Logging("=> Input reason")
            Commands.Wait10s_ClickElement(data["clock_out"]["confirm_clock_out"])
            Commands.Wait10s_ClickElement(data["clock_out"]["clock_out_success"])
            Logging("=> Clock out success")
        else:
            Logging("=> Clock out on time")
            Commands.Wait10s_ClickElement(data["clock_out"]["button_close"])
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
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        Commands.Wait10s_ClickElement("//*[@text='한국어']")
        Commands.Wait10s_ClickElement("//*[@text='닫기']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        
        korean_text = Waits.Wait10s_ElementInvisibility(data["settings"]["language_text"])
        if korean_text.text == '한국어':
            Logging("=> Change to language '한국어' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='언어 설정']")
        
        Commands.Wait10s_ClickElement("//*[@text='Tiếng Việt']")
        Commands.Wait10s_ClickElement("//*[@text='Đóng']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        VN_text = Waits.Wait10s_ElementInvisibility(data["settings"]["language_text"])
        if VN_text.text == 'Tiếng Việt':
            Logging("=> Change to language 'Tiếng Việt' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='Thay đổi ngôn ngữ']")

        Commands.Wait10s_ClickElement("//*[@text='日本語']")
        Commands.Wait10s_ClickElement("//*[@text='閉じる']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        JP = Waits.Wait10s_ElementInvisibility(data["settings"]["language_text"])
        if JP.text == '日本語':
            Logging("=> Change to language '日本語' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='言語']")

        Commands.Wait10s_ClickElement("//*[@text='简体中文']")
        Commands.Wait10s_ClickElement("//*[@text='關閉']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        TQ = Waits.Wait10s_ElementInvisibility(data["settings"]["language_text"])
        if TQ.text == '简体中文':
            Logging("=> Change to language '简体中文' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='語言']")

        Commands.Wait10s_ClickElement("//*[@text='Indonesian']")
        Commands.Wait10s_ClickElement("//*[@text='Tutup']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])

        indo = Waits.Wait10s_ElementInvisibility(data["settings"]["language_text"])
        if indo.text == 'Indonesian':
            Logging("=> Change to language 'Indonesian' success")
        else:
            Logging("=> Fail")

        Commands.Wait10s_ClickElement("//*[@text='Ganti BAHASA']")

        Commands.Wait10s_ClickElement("//*[@text='English']")
        Commands.Wait10s_ClickElement("//*[@text='Close']")
        Commands.Wait10s_ClickElement(data["menu_settings"]["button_settings"])
        EN = Waits.Wait10s_ElementInvisibility(data["settings"]["language_text"])
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
        approval_type = Waits.Wait10s_ElementInvisibility(data["event"]["popup"])
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
        Commands.Wait10s_ClickElement(data["menu_timecard"]["button_timecard"])
        Commands.Wait10s_ClickElement(data["vacation"]["button_vacation"])
        Logging("- Request Vacation")
        Commands.Wait10s_ClickElement(data["menu_timecard"]["request_vacation"])
        Logging("Save request")
    except:
        Logging("-> Can't request vacation")

def admin_settings_GPS():
    Commands.Wait10s_ClickElement(data["menu_admin"]["button_admin"])
    Logging("- Admin settings")

    Commands.Wait10s_ClickElement(data["menu_admin"]["GPS_setting"])
    Logging("- GPS Settings")
    Commands.Wait10s_ClickElement(data["menu_admin"]["add_GPS"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["popup"])
    Logging("- Add GPS")
    Commands.Wait10s_ClickElement(data["menu_admin"]["search_gps"])
    Logging("- Input GPS")
    Commands.InputElement(data["menu_admin"]["search_gps"], "Nguyen")
    Logging("- Enter value")
    Commands.Wait10s_ClickElement(data["menu_admin"]["done"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["gps"])
    Logging("- Select GPS")
    Commands.Wait10s_ClickElement(data["menu_admin"]["workplace"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["select_workplace"])
    Logging("- Select workplace")
    Commands.Wait10s_ClickElement(data["menu_admin"]["save_GPS"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["close_popup"])
    Logging("- Save GPS")

    driver.swipe(start_x=1000, start_y=450, end_x=500, end_y=450, duration=800)
    Commands.Wait10s_ClickElement(data["menu_admin"]["delete_gps"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["button_yes"])
    Logging("- Delete GPS")
    Commands.Wait10s_ClickElement(data["menu_admin"]["close_popup"])

def admin_settings_wifi():
    Commands.Wait10s_ClickElement(data["menu_admin"]["back_button"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["Wifi_setting"])
    Logging("- Wifi Settings")
    Commands.Wait10s_ClickElement(data["menu_admin"]["add_wifi"])
    Logging("- Add Wifi")

def admin_settings_beacon():
    Commands.Wait10s_ClickElement(data["menu_admin"]["back_button"])
    Commands.Wait10s_ClickElement(data["menu_admin"]["Beacon_setting"])
    Logging("- Beacon Settings")
    Commands.Wait10s_ClickElement(data["menu_admin"]["add_Beacon"])
    Logging("- Add Beacon")

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
        
    total_week_1 = Waits.Wait10s_ElementInvisibility(data["menu_timecard"]["list_tab"]["week_2_text"])
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

    total_week_2 = Waits.Wait10s_ElementInvisibility(data["menu_timecard"]["list_tab"]["week_3_text"])
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
    total_week_3 = Waits.Wait10s_ElementInvisibility(data["menu_timecard"]["list_tab"]["week_4_text"])
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
    total_week_4 = Waits.Wait10s_ElementInvisibility(data["menu_timecard"]["list_tab"]["week_5_text"])
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

        schedule = Waits.Wait10s_ElementInvisibility(data["TimeCard"]["report_monthly"]["schedule_working"])
        if schedule.text == 'Scheduled working day':
            count_day = Waits.Wait10s_ElementInvisibility(data["TimeCard"]["report_monthly"]["count_schedule_working"])
            Logging("- Scheduled working day:", count_day.text)
        else:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["report_monthly"]["events"])
        Logging("- Events")
        time.sleep(5)
        clock_in = Waits.Wait10s_ElementInvisibility(data["TimeCard"]["report_monthly"]["clockin"])
        if clock_in.text == 'Clock-In':
            count_clock_in = Waits.Wait10s_ElementInvisibility(data["TimeCard"]["report_monthly"]["count_clockin"])
            Logging("- Events - Clock in:", count_clock_in.text)
        else:
            Logging("=> Crash app")
            exit(0)
        time.sleep(5)

        Commands.Wait10s_ClickElement(data["TimeCard"]["report_monthly"]["working_status"])
        Logging("- Working status")
        time.sleep(5)
        working_time = Waits.Wait10s_ElementInvisibility(data["TimeCard"]["report_monthly"]["workingtime"])
        if working_time.text == 'Working time':
            count_working_time = Waits.Wait10s_ElementInvisibility(data["TimeCard"]["report_monthly"]["count_workingtime"])
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



Logging("Như Quỳnh")
execution()