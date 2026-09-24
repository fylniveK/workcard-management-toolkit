import os
import pyautogui
import time
from time import sleep
import openpyxl
import pyperclip
from sympy.plotting.intervalmath import interval

from general import *
from others import *
from mro import *
import subprocess
import webbrowser
import pygetwindow as gw


def document_cloud_auto_login(user_name:str, password:str):
    # 进入网页
    webbrowser.open(document_cloud_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle(config_constant['window_titles']['document_cloud'])
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦
    def judge_case(): # 判断当前在什么界面
        sleep(1)
        for i in range(20):
            pyautogui.hotkey('ctrl', '0', interval=0.1)  # 网页缩放至原大小
            flag1, _ = mouse_move_center(document_cloud_img_path, 'already_login.png')
            if flag1: # 说明在已登录界面
                return 1
            else: # 不在已登录界面
                flag2, _ = mouse_move_center(document_cloud_img_path, 'login_interface.png')
                flag3, _ = mouse_move_center(document_cloud_img_path, 'login_interface_150.png') # 缩放150的情况
                if flag2 or flag3: # 说明在登录界面
                    return 2
            sleep(0.1)

    def user_login():
        pyautogui.press('tab', presses=2, interval=0.1)
        sleep(1.5)
        paste_write(user_name)
        sleep(0.5)
        pyautogui.press('tab')
        paste_write(password)
        sleep(0.5)
        pyautogui.press('tab')
        # 多次尝试验证码，如果报错，报错图片存在5s左右
        try_ocr(document_cloud_img_path, 'captcha_error.png', airline_login_position_150, airline_login_captcha_enter_position)
        sleep(0.5)
        pyautogui.press('enter')

    def exit():
        for i in range(20):
            pyautogui.moveTo(1870,160,duration=0.2)
            sleep(1)
            flag,_ = mouse_move_center(document_cloud_img_path,'quit_account.png')
            pyautogui.click()
            if flag:
                sleep(0.2)
                pyautogui.click()
                break


    flag = judge_case()
    scale_to_150()
    if flag == 2:
        user_login()
    elif flag == 1:
        for i in range(5):
            exit()
            sleep(3)
            flag,_ = mouse_move_center(document_cloud_img_path,'login_interface_150.png',2)
            if flag:
                user_login()
                break

    sleep(2)
    pyautogui.hotkey('ctrl', '0', interval=0.1) # 网页缩放至原大小

if __name__ == '__main__':
    mro_user_name, mro_password, mro_mel_license = mro_account()
    document_cloud_auto_login(mro_user_name, mro_password)