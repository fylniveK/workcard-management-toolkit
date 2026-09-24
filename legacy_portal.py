import os
import pyautogui
import time
from time import sleep
import openpyxl
import pyperclip
from general import *
from others import *
import subprocess
import yaml
from datetime import datetime

current_script_dir = get_current_path() # 获取当前脚本所在位置（非执行位置）的绝对路径
legacy_portal_img_path = os.path.join(current_script_dir, "img", "LEGACY_PORTAL") # 拼接获取旧版维修平台图库绝对路径
employee = ''
work_order = ''
deliver_to = ''
signoff_employee = ''
signoff_password = ''
legacy_portal_title = config_constant['window_titles']['legacy_portal'] # 旧版维修平台的登录界面标题

# def legacy_portal_login():
#     '''
#     登录旧版维修平台函数，其中注意网页缩放为115%，改变缩放会导致本函数失效
#     :return: 无
#     '''
#     legacy_portal_url = 'https://example.invalid/legacy-portal/login'
#     legacy_portal_username = 'USER_C'
#     legacy_portal_password = 'REPLACE_WITH_PRIVATE_PASSWORD_C'
#     window_active(legacy_portal_title, sjzc_path)
#     pyautogui.hotkey('win','1',interval=0.1) # 这只是测试，后续要删掉
#     sleep(4)
#     paste_write(legacy_portal_url)
#     sleep(0.3)
#     pyautogui.press('enter') # 进入旧版维修平台网址
#     sleep(4)
#     pyautogui.hotkey('ctrl','w',interval=0.1) # 关闭旧版维修平台通知窗口
#     sleep(0.5)
#     paste_write(legacy_portal_username) # 输入签署人员旧版维修平台账号
#     sleep(0.5)
#     pyautogui.press('tab')
#     sleep(0.5)
#     paste_write(legacy_portal_password) # 输入签署人员旧版维修平台密码
#     sleep(0.3)
#     pyautogui.press('enter') # 登录
#     sleep(3)
#
#     # 登录账号后，点击绿色的勾
#     mouse_move_center(legacy_portal_img_path,'check.png',0.95)
#     pyautogui.click()
#     sleep(3)
#     # 关闭自动打开的另一个窗口
#     pyautogui.hotkey('alt','tab',interval=0.1)
#     pyautogui.press('enter')
#     pyautogui.hotkey('ctrl','w',interval=0.1)
#     sleep(0.5)
#     # 最大化旧版维修平台窗口
#     pyautogui.hotkey('win','up',interval=0.1)
#     sleep(0.3)
#
# def take_out_material(): # 进入旧版维修平台领料界面
#     # pyautogui.hotkey('win', '1', interval=0.1) # 这里可能要改一下，改为窗口激活而不是win+1
#
#     mouse_move_not_center(legacy_portal_img_path,'Navigation.png',0.95, 90, 30)
#     pyautogui.click()
#     sleep(0.5)
#
#     mouse_move_center(legacy_portal_img_path,'check_maintenance.png',0.9)
#     pyautogui.click()
#     sleep(3)
#
#     mouse_move_not_center(legacy_portal_img_path,'parts_request.png',0.9,9,5)
#     pyautogui.click()
#     sleep(1)
#
#     mouse_move_center(legacy_portal_img_path,'single_entry.png',0.9)
#     pyautogui.click()
#     sleep(1)
#
#     paste_write(employee)
#     pyautogui.press('tab',2)
#     sleep(0.5)
#
#     paste_write(work_order)
#     pyautogui.press('tab',5)
#     sleep(0.5)
#
#     paste_write(deliver_to)
#     pyautogui.press('tab',1)
#     sleep(0.5)
#
# def Solution_Signoff(): # 完成旧版维修平台最后的账号密码以及输入时间的部分
#     now = datetime.now()
#     date_str = now.strftime('%d-%m-%Y') # 日期24-10-2025格式
#     time_str = now.strftime('%H%M') # 时间0211格式（2点11分）
#     pyautogui.press('tab', 3)
#     paste_write(signoff_employee)
#     sleep(0.2)
#     pyautogui.press('tab')
#     sleep(0.1)
#     paste_write(signoff_password)
#     sleep(0.2)
#     pyautogui.press('tab', 3)
#     sleep(0.1)
#     paste_write(date_str)
#     sleep(0.2)
#     pyautogui.press('tab')
#     sleep(0.1)
#     paste_write(time_str)
#
# if __name__ == '__main__':
#     legacy_portal_login()
#     # sleep(1)
#
#     # take_out_material()
#
#     # Solution_Signoff()
