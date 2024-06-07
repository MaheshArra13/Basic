# import platform
# import re
# import subprocess
# from datetime import datetime
# from openpyxl import Workbook
# import time
# from appium.webdriver.common.appiumby import AppiumBy
#
# from framework_constants import Ios_franklin_Constants
#
#
# def get_current_time():
#     """
#     This function returns current time
#     :return:
#     """
#     current_time = datetime.now()
#     return current_time
#
# def execution_time(tc_start_time):
#     """
#         This function returns overall execution time
#         :return:
#     """
#     tc_end_time = get_current_time()
#     tc_time_interval = tc_end_time - tc_start_time
#     tc_time_interval_str = str(tc_time_interval).split(".")[0]
#     print("****************Execution completed***********************")
# #-------------------------- Installing .app for iOS Simulator -------------------
#
# app_path = "/Users/nadabala/Downloads/Convatec.app"
#
# def checking_app_installation(serial):
#     try:
#         cmd = f'xcrun simctl listapps {serial}'
#         process = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
#         apps = process.stdout.read().strip().decode()
#         if "FranklinDev" in apps:
#             print("Franklin application is already installed")
#         else:
#             print("")
#             print("Franklin application is not installed installing ..........")
#             install_cmd = f'xcrun simctl install {serial} {app_path}'
#             print("")
#             process_install = subprocess.Popen(install_cmd, stdout=subprocess.PIPE, shell=True)
#             cmd = f'xcrun simctl listapps {serial}'
#             process = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
#             print('Wait untill franklin application installation ....')
#             time.sleep(4)
#             apps = process.stdout.read().strip().decode()
#             if "com.convatec.franklindev" in apps:
#                 print("Franklin application is installed")
#     except Exception as error:
#         print(error)
#
#
#
# #--------------------------Checking appium launched or not----------------------------
#
#
# # Operating system
#
# WINDOWS = "Windows"
# iOS = "Darwin"
# def operating_system():
#     """
#     This funcrion returns the home directory of the machine
#     :return:
#     """
#     if platform.system() == iOS:
#         return iOS
#     elif platform.system() == WINDOWS:
#         return WINDOWS
#     else:
#         return platform.system()
#
# OPERATING_SYSTEM = operating_system()
#
#
#
#
#
# def Checking_appium():
#     cmd_windows = 'netstat -an | findstr /C:"4723"'
#     cmd_linux = "lsof -i :4723"
#     if OPERATING_SYSTEM == WINDOWS:
#         process_windows = subprocess.getoutput(cmd_windows)
#         if process_windows:
#             return True
#         else:
#             return False
#     elif OPERATING_SYSTEM == iOS:
#         proocess_linux = subprocess.getoutput(cmd_linux)
#         if proocess_linux:
#             print("")
#             print("Appium is running in port number : 4723")
#             print("")
#             return True
#         else:
#             print("Please check Appium is running in port number : 4723")
#             return False
#     else:
#         return False
#
#
#
#
# def get_iOS_devices():
#
#     ios_emulator_output = subprocess.Popen('xcrun simctl list devices | grep "(Booted)"', stdout=subprocess.PIPE,
#                                            shell=True)
#     output2 = ios_emulator_output.stdout.read().strip().decode()
#     ios_emulator_output.stdout.close()
#     ios_emulator = []
#     if output2:
#         # ios_emulator_output.terminate()
#         out = output2.split("\n")
#         for i in out:
#             dev_output1 = re.search(r"\b([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\b", i)
#             if dev_output1:
#                 dev_list1 = dev_output1.group(0)
#                 ios_emulator.append(dev_list1.strip())
#
#     return ios_emulator
#
#
#
# def desired_caps_IOS(serial):
#     from appium import webdriver
#     from appium.options.common import AppiumOptions
#     print("Serail is ::::",serial)
#     desired_caps = {}
#     desired_caps['platformName'] = "iOS"
#     desired_caps['deviceName'] = "iPhone 15 Plus"
#     desired_caps['udid'] = serial
#     desired_caps['bundleId'] = "com.convatec.franklindev"
#     desired_caps['platformVersion'] = "17.2"
#     desired_caps[''] = "XCUITest"
#     desired_caps['noReset'] = True
#     desired_caps['fullReset'] = False
#     desired_caps['showIOSLog'] = True
#     desired_caps['newCommandTimeout'] = 340
#     option = AppiumOptions().load_capabilities(desired_caps)
#     sender_driver = webdriver.Remote(command_executor=f"http://localhost:4723/wd/hub", options=option)
#     # sender_driver = webdriver.Remote('http://localhost:4723/wd/hub',desired_caps)
#     sender_driver.implicitly_wait(6)
#     print("Android Sender driver initialized")
#     return sender_driver
#
#
#
#
# def create_workbook():
#     """
#     This function create openpyxl workbook, test summary sheet and test report sheet objects
#     :param sheet1: sheet name of test report
#     :return: returns openpyxl workbook, test summary sheet and test report sheet objects
#     """
#     report_wokbook = Workbook()
#     summary_sheet = report_wokbook.active
#     summary_sheet.title = 'Test Summary'
#     # top_column = ("AZURE ID","Testcases Function","Status","Reason")
#     top_column = ("Testcases Function", "Status", "Reason")
#     summary_sheet.append(top_column)
#     return report_wokbook, summary_sheet
#
#
# def update_status(az_id, func_data, status,status_reason, sheet):
#     """
#     This function update the result in excel report
#     :param funcname: name of the funcion
#     :param func_data: dictionary of (testcase_data) pairs
#     :param status_reason: reason for execution started/not started
#     :param sheet: excel report test report
#     :return:
#     """
#     data = (
#         az_id, func_data,status, status_reason)
#     sheet.append(data)
#
#
import platform
###current
import re
import subprocess
from datetime import datetime
from openpyxl import Workbook
import time
from appium.webdriver.common.appiumby import AppiumBy

from framework_constants import *


def get_current_time():
    """
    This function returns current time
    :return:
    """
    current_time = datetime.now()
    return current_time

def execution_time(tc_start_time):
    """
        This function returns overall execution time
        :return:
    """
    tc_end_time = get_current_time()
    tc_time_interval = tc_end_time - tc_start_time
    tc_time_interval_str = str(tc_time_interval).split(".")[0]
    print("****************Execution completed***********************")
#-------------------------- Installing .app for iOS Simulator -------------------

app_path = "/Users/nadabala/Downloads/Convatec.app"

def checking_app_installation(serial):
    try:
        cmd = f'xcrun simctl listapps {serial}'
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
        apps = process.stdout.read().strip().decode()
        if "FranklinDev" in apps:
            print("Franklin application is already installed")
        else:
            print("")
            print("Franklin application is not installed installing ..........")
            install_cmd = f'xcrun simctl install {serial} {app_path}'
            print("")
            process_install = subprocess.Popen(install_cmd, stdout=subprocess.PIPE, shell=True)
            cmd = f'xcrun simctl listapps {serial}'
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
            print('Wait untill franklin application installation ....')
            time.sleep(4)
            apps = process.stdout.read().strip().decode()
            if "com.convatec.franklindev" in apps:
                print("Franklin application is installed")
    except Exception as error:
        print(error)



#--------------------------Checking appium launched or not----------------------------


# Operating system

WINDOWS = "Windows"
iOS = "Darwin"
def operating_system():
    """
    This funcrion returns the home directory of the machine
    :return:
    """
    if platform.system() == iOS:
        return iOS
    elif platform.system() == WINDOWS:
        return WINDOWS
    else:
        return platform.system()

OPERATING_SYSTEM = operating_system()





def Checking_appium():
    cmd_windows = 'netstat -an | findstr /C:"4723"'
    cmd_linux = "lsof -i :4723"
    if OPERATING_SYSTEM == WINDOWS:
        process_windows = subprocess.getoutput(cmd_windows)
        if process_windows:
            return True
        else:
            return False
    elif OPERATING_SYSTEM == iOS:
        proocess_linux = subprocess.getoutput(cmd_linux)
        if proocess_linux:
            print("")
            print("Appium is running in port number : 4723")
            print("")
            return True
        else:
            print("Please check Appium is running in port number : 4723")
            return False
    else:
        return False




def get_iOS_devices():

    ios_emulator_output = subprocess.Popen('xcrun simctl list devices | grep "(Booted)"', stdout=subprocess.PIPE,
                                           shell=True)
    output2 = ios_emulator_output.stdout.read().strip().decode()
    ios_emulator_output.stdout.close()
    ios_emulator = []
    if output2:
        # ios_emulator_output.terminate()
        out = output2.split("\n")
        for i in out:
            dev_output1 = re.search(r"\b([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\b", i)
            if dev_output1:
                dev_list1 = dev_output1.group(0)
                ios_emulator.append(dev_list1.strip())

    return ios_emulator



def desired_caps_IOS(serial):
    from appium import webdriver
    from appium.options.common import AppiumOptions
    print("Serail is ::::",serial)
    desired_caps = {}
    desired_caps['platformName'] = "iOS"
    desired_caps['deviceName'] = "iPhone 15 Plus"
    desired_caps['udid'] = serial
    desired_caps['bundleId'] = "com.convatec.franklindev"
    desired_caps['platformVersion'] = "17.2"
    desired_caps[''] = "XCUITest"
    desired_caps['noReset'] = True
    desired_caps['fullReset'] = False
    desired_caps['showIOSLog'] = True
    desired_caps['newCommandTimeout'] = 340
    option = AppiumOptions().load_capabilities(desired_caps)
    sender_driver = webdriver.Remote(command_executor=f"http://localhost:4723/wd/hub", options=option)
    print("DRIVER :" ,sender_driver )
    # sender_driver = webdriver.Remote('http://localhost:4723/wd/hub',desired_caps)
    sender_driver.implicitly_wait(6)
    print("Android Sender driver initialized")
    return sender_driver




def create_workbook():
    """
    This function create openpyxl workbook, test summary sheet and test report sheet objects
    :param sheet1: sheet name of test report
    :return: returns openpyxl workbook, test summary sheet and test report sheet objects
    """
    report_wokbook = Workbook()
    summary_sheet = report_wokbook.active
    summary_sheet.title = 'Test Summary'

    return report_wokbook, summary_sheet


def update_status(func_data, status,status_reason, sheet):
    """
    This function update the result in excel report
    :param funcname: name of the funcion
    :param func_data: dictionary of (testcase_data) pairs
    :param status_reason: reason for execution started/not started
    :param sheet: excel report test report
    :return:
    """
    data = (
        func_data,status, status_reason)
    sheet.append(data)


