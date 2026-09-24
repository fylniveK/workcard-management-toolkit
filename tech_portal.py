import os
import pyautogui
import time
from time import sleep
import pyperclip

from general import *
from others import *
from mro import *
import webbrowser
import pygetwindow as gw

def tech_portal_auto_login(user_name:str, password:str):
    webbrowser.open(tech_portal_login_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle(config_constant['window_titles']['tech_portal'])
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦
    def judge_case(): # 判断当前在什么界面
        sleep(1)
        for i in range(20):
            pyautogui.hotkey('ctrl', '0', interval=0.1)  # 网页缩放至原大小
            pyautogui.moveTo(*tech_portal_login_captcha_enter_position_150)
            flag1,_ = mouse_move_center(tech_portal_img_path,'login_entry.png')
            if flag1: # 说明在登陆界面
                return 1
            else:
                flag3, _ = mouse_move_center(tech_portal_img_path, 'already_login.png')
                if flag3:  # 说明已登录；也有可能是“去登陆”界面，因为这俩图同时出现，下面需要优先判断是否需要重新登陆
                    flag2, _ = mouse_move_center(tech_portal_img_path, 'go_to_login_button.png')
                    if flag2:  # 如果有这个图，说明现在登录状态失效，“去登陆”的界面
                        return 2
                    else:
                        return 3

    def user_login():
        pyautogui.moveTo(*tech_portal_login_captcha_enter_position_150)
        pyautogui.click()
        sleep(0.01)
        pyautogui.press('tab')
        paste_write(user_name)
        sleep(0.5)
        pyautogui.press('tab')
        pyautogui.hotkey('ctrl','a',interval=0.1)
        paste_write(password)
        sleep(0.5)
        pyautogui.press('tab')
        scale_to_150()
        try_ocr(tech_portal_img_path, 'captcha_error_150.png', tech_portal_login_position_150, tech_portal_login_captcha_enter_position_150)
        sleep(0.5)
        pyautogui.press('enter')
        sleep(0.5)
        pyautogui.hotkey('ctrl','0',interval=0.1)

    def exit():
        for i in range(20):
            flag,_ = mouse_move_center(tech_portal_img_path,'exit_button.png')
            if flag:
                sleep(0.1)
                pyautogui.click()
                break

    flag = judge_case()
    if flag == 1:
        user_login()
    elif flag == 3:
        for i in range(5):
            exit()
            sleep(3)
            pyautogui.moveTo(*tech_portal_login_captcha_enter_position_150)
            flag,_ = mouse_move_center(tech_portal_img_path,'login_entry.png',2)
            if flag:
                user_login()
                break
    elif flag == 2:
        for i in range(5):
            flag,_ = mouse_move_center(tech_portal_img_path,'go_to_login_button.png')
            if flag:
                pyautogui.click()
                sleep(2)
                pyautogui.moveTo(*tech_portal_login_captcha_enter_position_150)
                flag, _ = mouse_move_center(tech_portal_img_path, 'login_entry.png', 2)
                if flag:
                    user_login()
                    break

def open_manual_url():
    webbrowser.open(tech_portal_manual_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle(config_constant['window_titles']['tech_portal'])
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦
    sleep(2)

def tech_portal_737ng_pp_auto_open():
    open_manual_url()
    sleep(3)
    for i in range(5):
        flag,_ = mouse_move_center(tech_portal_img_path,'manual_boeing.png')
        if flag:
            pyautogui.click()
            break
    for i in range(5):
        flag, _ = mouse_move_center(tech_portal_img_path, 'manual_737ng.png')
        if flag:
            pyautogui.click()
            break
    for i in range(5):
        sleep(1)
        flag, _ = mouse_move_center(tech_portal_img_path, 'manual_737ng_pp.png')
        if flag:
            pyautogui.click()
            break

def tech_portal_737ng_ipc_auto_open():
    open_manual_url()
    sleep(3)
    for i in range(5):
        flag,_ = mouse_move_center(tech_portal_img_path,'manual_boeing.png')
        if flag:
            pyautogui.click()
            break
    for i in range(5):
        flag, _ = mouse_move_center(tech_portal_img_path, 'manual_737ng.png')
        if flag:
            pyautogui.click()
            break
    for i in range(5):
        sleep(1)
        flag, _ = mouse_move_center(tech_portal_img_path, 'manual_737ng_ipc.png')
        if flag:
            pyautogui.click()
            break


if __name__ == '__main__':
    mro_user_name, mro_password, mro_mel_license = mro_account()
    tech_portal_auto_login(mro_user_name,mro_password)
    # tech_portal_737ng_ipc_auto_open()