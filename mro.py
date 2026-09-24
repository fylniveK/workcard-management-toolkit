import os
import pyautogui
import time
from time import sleep
import openpyxl
import pyperclip
from general import *
from others import *
import subprocess
import webbrowser
import pygetwindow as gw


def mro_auto_login(user_name:str, password:str): # 在文档云登录的前提下，由于登录状态是通的，所以可以直接进mro的登录状态
    # 进入网页
    webbrowser.open(mro_login_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle(config_constant['window_titles']['mro_portal'])
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦

    def judge_case(): # 判断当前在什么界面
        sleep(1)
        for i in range(20):
            pyautogui.hotkey('ctrl', '0', interval=0.1)  # 网页缩放至原大小
            flag1, _ = mouse_move_center(mro_img_path, 'internal_employee_login.png')
            if flag1:  # 在欢迎界面
                return 1
            else: # 不在欢迎界面
                flag2, _ = mouse_move_center(mro_img_path, 'already_login.png')
                if flag2: # 在已登录界面
                    return 2
                else: # 不在已登录界面
                    flag3, _ = mouse_move_center(mro_img_path, 'login_time_out.png')
                    if flag3: # 有重新登录框弹出的界面
                        return 3
            sleep(0.1)

    flag = judge_case()
    if flag == 1: # 欢迎界面
        mouse_move_center(mro_img_path, 'internal_employee_login.png', 2)
        pyautogui.click()
    elif flag == 2: # 成功登录
        return
    elif flag == 3: # 有重新登录框弹出的界面
        mouse_move_center(mro_img_path, 'confirm_button.png', 2) # 找到确认按钮
        pyautogui.click()
        flag = judge_case()
        if flag == 1: # 欢迎界面
            mouse_move_center(mro_img_path, 'internal_employee_login.png', 2) # 找内部员工登录按钮
            sleep(0.1)
            pyautogui.click()
            flag = judge_case()
            if flag == 2: # 成功登录
                return

def mro_auto_query_aircraft_parts():
    # 进入网页
    webbrowser.open(mro_search_aircraft_parts_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle(config_constant['window_titles']['mro_portal'])
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦
    sleep(0.5)
    pyautogui.hotkey('ctrl','0',interval=0.1)

def mro_auto_tasks_issue(aircraft_id):
    # 进入网页
    webbrowser.open(mro_tasks_issue_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle(config_constant['window_titles']['mro_portal'])
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦
    sleep(0.5)
    pyautogui.hotkey('ctrl', '0', interval=0.1)

    for i in range(20): # 进入A检工作包界面
        flag1,_ = mouse_move_center(mro_img_path,'A_check_workcards_button.png') # 找到“A检工作包”按钮
        if flag1:
            pyautogui.click()
        flag2,_ = mouse_move_center(mro_img_path,'A_check_workcards_mark.png') # 看是否已经进入A检工卡包
        if flag2: # 如果找到了进入该界面的标志
            break
        sleep(0.1)
    sleep(0.5)
    mouse_move_center(mro_img_path,'flight_no_window.png', 4) # 找到飞机号输入框
    pyautogui.click()
    sleep(0.05)
    paste_write(aircraft_id)
    sleep(0.1)
    mouse_move_center(mro_img_path,'search_button.png', 4) # 找到查询
    pyautogui.click()

def mro_auto_non_routine_workcard_search(aircraft_id):
    # 进入网页
    webbrowser.open(mro_non_routine_workcard_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle(config_constant['window_titles']['mro_portal'])
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦
    sleep(0.5)
    pyautogui.hotkey('ctrl', '0', interval=0.1)

    flag1,_ = mouse_move_center(mro_img_path,'non_routine_workcard_ui.png', 4) # 看是否进入了非卡方案页面
    if not flag1: # 一直没找到的话
        print('长时间未进入非卡界面')
        return
    sleep(2)
    mouse_move_center(mro_img_path, 'unfold_button.png', 2) # 点击展开按钮
    pyautogui.click()
    sleep(0.3)
    mouse_move_center(mro_img_path, 'condition_search.png', 2)  # 点击条件搜索框
    pyautogui.click()
    sleep(0.3)
    paste_write(aircraft_id)
    # 获取昨天与明天的日期
    yesterday_str = get_time_standard('yesterday')
    today_str = get_time_standard('today')
    mouse_move_center(mro_img_path, 'cards_create_date.png', 2)  # 点击开始日期框
    pyautogui.click()
    sleep(0.5)
    paste_write(yesterday_str)
    sleep(1)
    pyautogui.press('tab')
    sleep(1)
    paste_write(today_str)
    sleep(0.5)
    mouse_move_center(mro_img_path, 'plan_condition.png', 2)  # 点击方案状态，这个放后面是因为网页会抽风，不出现未完成按钮
    pyautogui.click()
    sleep(1.5)
    flag2,_ = mouse_move_center(mro_img_path, 'uncompleted_button.png', 2)  # 点击未完成
    sleep(1)
    pyautogui.click()
    sleep(0.3)
    mouse_move_center(mro_img_path, 'search_button.png', 2)  # 点击查询按钮
    pyautogui.click()



# def mro_auto_login(user_name:str, password:str): # 这个逻辑过于复杂，目前好像可以自动登录，但是用时非常长，不好修改，最关键的问题是官方mro网站有bug，导致登录无权限账号后无法退出重新登录
#     # 进入网页
#     webbrowser.open(mro_login_url)
#     sleep(2)
#     # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
#     windows = gw.getWindowsWithTitle('TECH_PORTAL')
#     if windows:
#         browser = windows[0] # 打开其中第一个，不过这里大概率只有一个
#         browser.maximize() # 最大化
#         browser.activate() # 聚焦
#     sleep(1)
#
#     def welcome_login_case():# 1.是否有“内部员工登录”按钮，对应欢迎登录界面，需要点击内部员工登录按钮
#         flag, img_path = mouse_move_center(mro_img_path, 'internal_employee_login.png') # 此时鼠标已经放在对应图片位置
#         pyautogui.click()  # 点击进去了
#         sleep(1)
#         return flag, img_path
#     def already_login_case():# 2.是否有“already_login”图片，对应已登录界面，但不一定是我们想要登陆的账号
#         flag, img_path = mouse_move_center(mro_img_path, 'already_login.png')
#         return flag, img_path
#     def login_entry_case():# 3.是否有“login_entry”图片，对应实际官方登录入口界面，需要实际进行账号密码登录
#         flag, img_path = mouse_move_center(mro_img_path, 'login_entry.png')
#         return flag, img_path
#     def user_judge(tries: int=5, interval: float=0.2):# 判断当前登录用户是谁
#         tries = max(5,int(tries))
#         for i in range(tries):
#             current_user_no = ocr(mro_user_position)
#             if current_user_no == user_name:
#                 flag = True
#                 break
#             else:
#                 flag = False
#         return flag
#
#     def mro_login_state(max_time=10, interval: float=0.1): # 利用图片判断登录状态
#         state = False
#         tries = max(3, int(max_time / interval))
#         for i in range(tries):
#             # 1.是否有“内部员工登录”按钮，对应欢迎登录界面，需要点击内部员工登录按钮
#             flag1, img_path1 = welcome_login_case()
#             sleep(interval)
#             # 2.是否有“already_login”图片，对应已登录界面，但不一定是我们想要登陆的账号
#             flag2, img_path2 = already_login_case()
#             sleep(interval)
#             # 3.是否有“login_entry”图片，对应实际官方登录入口界面，需要实际进行账号密码登录
#             flag3, img_path3 = login_entry_case()
#             sleep(interval)
#             if flag1:
#                 state = 1
#                 print('当前在欢迎界面')
#             elif flag2:
#                 state = 2
#                 print('当前已登录')
#             elif flag3:
#                 state = 3
#                 print('当前在入口界面')
#
#             if state:
#                 return state
#             else:
#                 print('等待网页加载')
#         return state
#
#     def entry_login(): # 在官方登录入口界面进行登录
#         pyautogui.hotkey('ctrl','0', interval=0.1) # 网页缩放至原大小
#         pyautogui.press('tab', presses=2, interval=0.1)
#         sleep(1.5)
#         paste_write(user_name)
#         pyautogui.press('tab')
#         paste_write(password)
#         for i in range(3):
#             pyautogui.hotkey('ctrl','+',interval=0.1)
#             sleep(0.2)
#         pyautogui.press('tab')
#         try_ocr(mro_img_path, 'captcha_error.png', airline_login_position_150) # 多次尝试验证码，如果报错，报错图片存在5s左右
#         pyautogui.hotkey('ctrl', '0', interval=0.1) # 网页缩放至原大小
#
#     def quit(): # 退出登录
#         pyautogui.press('tab', presses=3, interval=0.2)
#         sleep(0.5)
#         pyautogui.press('enter')
#         sleep(0.5)
#         pyautogui.press('down', presses=3, interval=0.2)
#         sleep(0.5)
#         pyautogui.press('enter')
#         sleep(1)
#
#     def re_login(): # 重新登录完整流程，在确定已登录且号不对的情况下
#         quit()
#         flag,_ = welcome_login_case()
#         sleep(0.1)
#         pyautogui.click()
#         sleep(1)
#         entry_login()
#
#     def judge_and_login(): # 判断已登录账号用户名是否正确，不正确就重登
#         flag = user_judge()
#         if not flag:
#             re_login()
#         return flag
#
#
#     # 正式开始完整登录流程
#     state = mro_login_state() # 判断状态
#     if state == 1: # 欢迎界面
#         state = mro_login_state()
#         if state == 2: # 已登录界面
#             # 判断账号是否是我需要的
#             flag = judge_and_login() # 如果不是就已经重登了
#             if flag: # 如果是的
#                 return # 结束任务
#         elif state == 3: # 官方实际登录入口
#             entry_login()
#             # 判断账号是否是我需要的
#             flag = judge_and_login() # 如果不是就已经重登了
#             if flag:  # 如果是的
#                 return  # 结束任务
#     elif state == 2: # 已登录界面
#         # 判断账号是否是我需要的
#         flag = judge_and_login() # 如果不是就已经重登了
#         if flag:  # 如果是的
#             return  # 结束任务




if __name__ == '__main__':
    # mro_user_name, mro_password, mro_mel_license = mro_account()
    # mro_auto_login(mro_user_name, mro_password)
    # mro_auto_query_aircraft_parts()
    pass