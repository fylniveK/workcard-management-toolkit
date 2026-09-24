import os
import pyautogui

from general import *
import tkinter as tk
import webbrowser
import pygetwindow as gw

# 当前脚本目录
current_script_dir = get_current_path()
# 桌面目录
desktop_path = config_constant['paths']['desktop']
# 脚本内other_files_and_models文件夹目录
other_files_and_models_path = os.path.join(current_script_dir, 'other_files_and_models')
# 图片总路径
general_img_path = os.path.join(current_script_dir, 'img')

sjzc_path = config_constant['paths']['alternate_browser']
chrome_path = config_constant['paths']['chrome']

XiaoYeDan_file_path_60 = os.path.join(other_files_and_models_path, '夜宵（60）.docx')
XiaoYeDan_file_path_50 = os.path.join(other_files_and_models_path, '夜宵（50）.doc')
XiaoYeDan_file_path_40 = os.path.join(other_files_and_models_path, '夜宵（40）.doc')
ChangYongJianHao_file_path = config_constant['paths']['common_parts']
Local_IPC_dir_path = config_constant['paths']['local_ipc']
HangCaiXuQiuDan_template_path = os.path.join(other_files_and_models_path, '航材需求单模板.xlsx')
JiaoJieDan_template_path = os.path.join(other_files_and_models_path, '定检工作交接单模板.docx')
UseGuide_url = config_constant['links']['use_guide']
GongKaYuanZhiNan_url = config_constant['links']['workcard_guide']
ChaJian_JianHao_url = config_constant['links']['parts_guide']
CaoZuoZhiNan_url = config_constant['links']['operation_guide']
BaoLiuGuZhangKongZhiDan_template_path = os.path.join(other_files_and_models_path, 'EP02-008F01保留故障控制单0712版模板.doc')
ng_ShiCheDan_file_path = os.path.join(other_files_and_models_path, 'CFM、LEAP试车报告单_7jEfMf0440588.pdf')
C909_ChangYongJianHao_file_path = os.path.join(other_files_and_models_path, '909常用航材.docx')
C909_ShiCheDan_file_path = os.path.join(other_files_and_models_path,'909试车报告单.docx')
auto_man_hours_template_path = config_constant['paths']['man_hours_template']
auto_man_hours_fetch_path = config_constant['paths']['man_hours_fetch']
auto_man_hours_document_path = config_constant['paths']['man_hours_document']
KeCangBiaoPai_file_path = os.path.join(other_files_and_models_path,'客舱标牌.pdf')
# 上外网登录网址
go_online_url = config_constant['links']['network_login']

# 拿到others文件夹中的图片总路径
others_img_path = os.path.join(general_img_path,'others')
online_login_position_150 = (1013, 546, 180, 61) # 登录界面缩放150%时验证码位置信息
# 如果验证码错误，位置会改变
online_login_position_change_150 = (1013, 591, 180, 61) # 登录界面缩放150%时验证码位置信息
online_login_captcha_enter_position_change = (864, 622)

# document_cloud所需路径
document_cloud_img_path = os.path.join(general_img_path,'document_cloud')
document_cloud_url = config_constant['links']['document_cloud']
man_hours_collect_url = config_constant['links']['man_hours']
quality_rectification_url = config_constant['links']['quality_rectification']
airline_login_position_150 = (1638, 550, 165, 64) # 登录界面缩放150%时验证码位置信息
airline_login_captcha_enter_position = (1338, 593)

# mro所需路径
mro_img_path = os.path.join(general_img_path,'MRO')
mro_login_url = config_constant['links']['mro_login']
mro_user_position = (324, 313, 102, 28) # 登录上去之后员工号信息位置，用于判断目前账号内容
mro_non_routine_workcard_url = config_constant['links']['mro_non_routine_workcard']
mro_tasks_issue_url = config_constant['links']['mro_tasks_issue']
mro_search_aircraft_parts_url = config_constant['links']['mro_parts_search']

# tech_portal所需路径
tech_portal_img_path = os.path.join(general_img_path,'TECH_PORTAL')
tech_portal_login_url = config_constant['links']['tech_portal_login']
tech_portal_manual_url = config_constant['links']['tech_portal_manual']
tech_portal_login_position_150 = (1005, 832, 165, 62) # 登录界面缩放150%时验证码位置信息
tech_portal_login_captcha_enter_position_150 = (780, 872)

def mro_account(name:str = 'operator_b'): # 获取mro账号密码信息
    account = config_constant['mro']['account'][name]
    user_name = account['user_name']
    password = account['password']
    mel_license = account['mel_license']
    return user_name, password, mel_license

def login_online(user_name:str, password:str):
    webbrowser.open(go_online_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle('上网认证系统')
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦

    def judge_case():
        for i in range(5):
            flag1, _ = mouse_move_center(others_img_path, 'online_login_entry_150.png')
            if flag1:
                return 1 # 登录界面
            sleep(0.05)
            flag2, _ = mouse_move_center(others_img_path,'online_ok_150.png')
            if flag2:
                return 2 # 登录成功界面

    def try_ocr(current_img_path: str, error_img_name: str, ocr_position: tuple, captcha_position: tuple, ocr_position_change: tuple,
                login_button_name: str = '', confidence: float = 0.9, tries: int = 10, interval: float = 0.2):
        """
        这里基本照搬general中的函数，但是这里登录比较特殊所以需要特别调整
        """
        # 这里的captcha_position是为了拿到验证码输入的位置，方便重输
        error_img_path = os.path.join(current_img_path, error_img_name)
        x,y,_,_ = ocr_position_change
        # 为了拿到最后验证码中间的位置
        x += 84
        y += 26
        # 第一次尝试，位置在原位
        captcha = ocr(ocr_position)  # 获取验证码信息
        paste_write(captcha)  # 输入验证码
        sleep(interval)
        mouse_move_center(current_img_path, login_button_name) # 移动到登录按钮处
        pyautogui.click()
        sleep(interval)
        # 后续如果失败持续尝试，整个位置改变
        for i in range(tries):
            captcha = ocr(ocr_position_change)  # 获取验证码信息
            paste_write(captcha)  # 输入验证码
            sleep(interval)
            mouse_move_center(current_img_path, login_button_name)  # 移动到登录按钮处
            pyautogui.click()
            sleep(interval)
            # 判断是否有验证码报错
            flag = bool(find_img_center(error_img_path, confidence=confidence, max_loops=tries, interval=interval))
            if not flag: # 如果没找到那就说明成功登录，退出循环
                break
            sleep(0.5)
            pyautogui.click(x,y)
            sleep(0.3)
            pyautogui.doubleClick(*captcha_position) # 选中验证码区域，这里*captcha_position是解包元组，把元组中俩值传给x和y
            sleep(0.1)
            pyautogui.hotkey('ctrl', 'a', interval=0.1) # 快捷键全选，确保一定全选了验证码，方便后面输入新的
            sleep(0.5)

    def user_login():
        paste_write(user_name)
        sleep(0.5)
        pyautogui.press('tab')
        paste_write(password)
        sleep(0.5)
        pyautogui.press('tab')
        try_ocr(others_img_path,'online_captcha_error_150.png',online_login_position_150,online_login_captcha_enter_position_change,
                online_login_position_change_150, 'online_login_button_150.png')

    scale_to_150()
    flag = judge_case()
    if flag == 1:
        user_login()
    elif flag == 2:
        return



if __name__ == '__main__':
    mro_user_name, mro_password, mro_mel_license = mro_account()
    login_online(mro_user_name, mro_password)
