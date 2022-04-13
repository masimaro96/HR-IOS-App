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

APPIUM_PORT = '4723'
udid = 'bc86e429485c13f34837866fde36e7ed55646317'
app_path = 'Users/hanbiro/Desktop/nhuquynhios/HanbiroHR.ipa'
command_executor ='http://127.0.0.1:%s/wd/hub' % APPIUM_PORT

desired_capabilities = {
    'orientation' :'LANDSCAPE',
    "deviceName": "Hanbiro Iphone",
    "platformVersion": "12.5.5",
    "platformName": "IOS",
    "udid": udid,
    "app": app_path,
    "autoWebView":True,
    "fullContextList" : True,
    "enablePerformanceLogging" :True
}

driver = webdriver.Remote(command_executor,desired_capabilities)
with open("NQ_HR.json") as json_data_file:
    data = json.load(json_data_file)

n = random.randint(1,3000)

class objects:
    now = datetime.now()
    year = now.strftime("%Y")
    month = now.strftime("%m")
    day = now.strftime("%d")
    time1 = now.strftime("%H:%M:%S")
    date_time = now.strftime("%Y/%m/%d, %H:%M:%S")
    date_id = date_time.replace("/", "").replace(", ", "").replace(":", "")[2:]
    testcase_pass = "Test case status: pass"
    testcase_fail = "Test case status: fail"

if platform == "linux" or platform == "linux2":
    local_path = "/home/oem/groupware-auto-test"
    json_file = local_path + "/NQ_HR.json"
    with open(json_file) as json_data_file:
        data = json.load(json_data_file)
    log_folder = "/Log/"
    log_testcase = "/Log/"
    execution_log = local_path + log_folder + "hanbiro_HR_execution_log_" + str(objects.date_id) + ".txt"
    fail_log = execution_log.replace("hanbiro_HR_execution_log_", "fail_log_")
    error_log = execution_log.replace("hanbiro_HR_execution_log_", "error_log_")
    testcase_log = local_path + log_testcase + "NQuynh_Testcase_HRApp_" + str(objects.date_id) + ".xlsx"
else :
    local_path = "/Users/hanbiro/Desktop/luuios"
    json_file = local_path + "/NQ_HR.json"
    with open(json_file) as json_data_file:
        data = json.load(json_data_file)
    log_folder = "/Log/"
    log_testcase = "/Log/"
    execution_log = local_path + log_folder + "hanbiro_HR_execution_log_" + str(objects.date_id) + ".txt"
    fail_log = execution_log.replace("hanbiro_HR_execution_log_", "fail_log_")
    error_log = execution_log.replace("hanbiro_HR_execution_log_", "error_log_")
    testcase_log = local_path + log_testcase + "NQuynh_Testcase_HRApp_" + str(objects.date_id) + ".xlsx"

logs = [execution_log, fail_log, error_log, testcase_log]
for log in logs:
    if ".txt" in log:
        open(log, "x").close()
    else:
        wb = Workbook()
        myFill = PatternFill(start_color='adc5e7',
                   end_color='adc5e7',
                   fill_type='solid',)
        font = Font(name='Calibri',
                    size=11 ,
                    bold=True,
                    italic=False,
                    vertAlign=None,
                    underline='none',
                    strike=False,
                    color='FF000000')
        ws = wb.active

        ws.cell(row=1, column=1).value= "Menu"
        ws.cell(row=1, column=2).value = "Sub-Menu"
        ws.cell(row=1, column=3).value = "Test Case Name"
        ws.cell(row=1, column=4).value = "Status"
        ws.cell(row=1, column=5).value = "Description"
        ws.cell(row=1, column=6).value = "Date"
        ws.cell(row=1, column=7).value = "Tester"
        # color 
        ws.cell(row=1, column=1).fill = myFill
        ws.cell(row=1, column=2).fill = myFill
        ws.cell(row=1, column=3).fill = myFill
        ws.cell(row=1, column=4).fill = myFill
        ws.cell(row=1, column=5).fill = myFill
        ws.cell(row=1, column=6).fill = myFill
        ws.cell(row=1, column=7).fill = myFill
        # font
        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=2).font = Font(bold=True)
        ws.cell(row=1, column=3).font = Font(bold=True)
        ws.cell(row=1, column=4).font = Font(bold=True)
        ws.cell(row=1, column=5).font = Font(bold=True)
        ws.cell(row=1, column=6).font = Font(bold=True)
        ws.cell(row=1, column=7).font = Font(bold=True)

        wb.save(log)

def Logging(msg):
    print(msg)
    log_msg = open(execution_log, "a")
    log_msg.write(str(msg) + "\n")
    log_msg.close()

def TesCase_LogResult(menu, sub_menu, testcase, status, description, tester):
    Logging(description)

    # if status == "Pass":
    #     Logging(objects.testcase_pass)
    # else:
    #     Logging(objects.testcase_fail)

    wb = openpyxl.load_workbook(testcase_log)
    current_sheet = wb.active
    start_row = len(list(current_sheet.rows)) + 1

    current_sheet.cell(row=start_row, column=1).value = menu
    current_sheet.cell(row=start_row, column=2).value = sub_menu
    current_sheet.cell(row=start_row, column=3).value = testcase
    current_sheet.cell(row=start_row, column=4).value = status
    current_sheet.cell(row=start_row, column=5).value = description
    current_sheet.cell(row=start_row, column=6).value = objects.date_time
    current_sheet.cell(row=start_row, column=7).value = tester

    # Apply color for status: Pass/Fail
    passFill = PatternFill(start_color='b6d7a8',
                   end_color='b6d7a8',
                   fill_type='solid',)
    failFill = PatternFill(start_color='ea9999',
                   end_color='ea9999',
                   fill_type='solid')
    if status == "Pass":
        Logging(objects.testcase_pass)
        current_sheet.cell(row=start_row, column=4).fill = passFill
    else:
        Logging(objects.testcase_fail)
        current_sheet.cell(row=start_row, column=4).fill = failFill
    wb.save(testcase_log)

def ValidateFailResultAndSystem(fail_msg):
    Logging(fail_msg)
    append_fail_result = open(fail_log, "a")
    append_fail_result.write("[FAILED TEST CASE] " + str(fail_msg) + "\n")
    append_fail_result.close()

def execution():
    Logging("------- Login to app -------")
    # Input information for log-in
    driver.find_element_by_ios_class_chain(data["domain"]).send_keys(data["domain_name"])
    Logging("- Input Domain")
    driver.find_element_by_ios_class_chain(data["id_login"]).send_keys(data["id_login_name_1"])
    Logging("- Input ID")
    driver.find_element_by_ios_class_chain(data["pass_login"]).send_keys(data["password"])
    Logging("- Input Password")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["login_button"]))).click()
    Logging("=> Click Log In button")
    driver.implicitly_wait(50)


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

def add_event():
    try:
        Logging(" ")
        Logging("------- Add event -------")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["timecard"]))).click()
        Logging("- Select time card")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["timesheet"]))).click()
        Logging("- Select time sheet")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["add"]))).click()
        Logging("- Select add")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@text,'Please input data.')]"))).send_keys(data["event"]["title_text"])
        Logging("- Input title")
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["choose_event"]))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["type_event"]))).click()
        Logging("- Choose event")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["select_color"]))).click()
        Logging("- Select color")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@text,'Please input data.')]"))).send_keys(data["event"]["location_text"])
        Logging("- Input location")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@text,'Please add a memo.')]"))).send_keys(data["event"]["memo_text"])
        Logging("- Input memo")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["event"]["button_save"]))).click()
        TesCase_LogResult(**data["testcase_result"]["timecard"]["event"]["pass"])
    except:
        Logging("- Can't create event")
        TesCase_LogResult(**data["testcase_result"]["timecard"]["event"]["fail"])

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
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["button_vacation"]))).click()
    print("- Vacation")
    ''' Request vaction'''
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["request"]))).click() 
    
    title_request = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["request_vacation_text"])))
    if title_request.text == 'Request vacation':
        print("=> Request vacation")
    else:
        print("=> Crash app")
        exit(0)   
    try:
        vacation_type = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["AM"])))
        if vacation_type.is_displayed():
            vacation_type.click()
            print("- Select vacation type")

            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["calendar"]))).click()
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]"))).click()
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]"))).click()
            time.sleep(2)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["select_calendar"]))).click()  
            time.sleep(2)

            ''' Crash app when select date '''
            try:
                title_request = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["request_vacation_text"])))
                if title_request.text == 'Request vacation':
                    print("- Select date vacation")
                else:
                    print("=> Crash app")
                    exit(0)  
            except WebDriverException: 
                print("=> Crash app")
                exit(0)

            ''' Get data of vacation request '''
            vacation_date = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["request_date_text"])))
            vacation_text = vacation_date.text
            date_text = vacation_text.split(" ")[0]
            type_vacation = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["AM"])))
            type_text = type_vacation.text
            vacation_date_type = date_text + "(" + type_text + ")"            
    except:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["calendar"]))).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]"))).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]"))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["select_calendar"]))).click()  
        time.sleep(2)

        ''' Get data of vacation request '''
        vacation_date = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["request_date_text"])))
        vacation_text = vacation_date.text
        # date_text = vacation_text.split(" ")[0]
        # type_vacation = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["AM"])))
        # type_text = type_vacation.text
        # vacation_date_type = vacation_text + "(" + type_text + ")"

    driver.swipe(start_x=650, start_y=1844, end_x=650, end_y=355, duration=800)
    time.sleep(5)
    ''' Select CC '''
    # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["CC"]))).click()
    # input_user_name()
    # WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@text,'quynh1')]"))).click()
    # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["save_cc"]))).click()
    # print("- Select CC")

    driver.swipe(start_x=650, start_y=1662, end_x=650, end_y=355, duration=800)

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@text='Please enter your reason']"))).send_keys(data["vacation"]["my_vacation"]["input_test"])
    print("- Input reason")
    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["button_request"]))).click()
    
    '''- Check day request
      + If vacation day is saturday => fail, check again
      + If memo is empty => fail, check again'''
    try:
        fail = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["request_fail"])))
        if fail.text == 'request vacation failure':
            print("--- Request vacation failure - vacation day is saturday---")
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["close_fail"]))).click()
            time.sleep(2)
            driver.swipe(start_x=650, start_y=355, end_x=650, end_y=2275, duration=800)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["calendar"]))).click()
            time.sleep(2)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]"))).click()
            print("=> Select start date")
            
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//android.view.ViewGroup[@index='1']//android.widget.Button[@index=16]"))).click()
            print("=> Select end date")
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["select_calendar"]))).click()
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["button_request"]))).click() 
            print("=> Send request vacation")
            TesCase_LogResult(**data["testcase_result"]["vacation"]["request_vacation"]["pass"])
        else:
            print("=> Request success")
            TesCase_LogResult(**data["testcase_result"]["vacation"]["request_vacation"]["pass"])
    except WebDriverException:
        print("=> Request success") 
        TesCase_LogResult(**data["testcase_result"]["vacation"]["request_vacation"]["pass"])

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["vacation"]["my_vacation"]["button_close"]))).click()

def cancel_request():
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["cancel_request"]))).click()
        print("- Cancel request")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status"]))).click()

        status_option1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status_request"])))
        status_name1 = str(status_option1.text)
        status_option1.click()
        print("- Status: " + status_name1)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status"]))).click()
        status_option2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status_processing"])))
        status_name2 = str(status_option2.text)
        status_option2.click()
        print("- Status: " + status_name2)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status"]))).click()
        status_option3 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status_waiting"])))
        status_name3 = str(status_option3.text)
        status_option3.click()
        print("- Status: " + status_name3)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status"]))).click()
        status_option4 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["access_pages_vacation"]["vacation_approve"]["status_complete"])))
        status_name4 = str(status_option4.text)
        status_option4.click()
        print("- Status: " + status_name4)
    except:
        pass
    
#def view_noti():

# def settings():
#     try:
#         Logging("--- Settings - Change language ---")
#         WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["menu_settings"]["button_settings"]))).click()
        
#     except WebDriverException:
#         Logging("=> Crash app")

