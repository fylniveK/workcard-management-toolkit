import os
import time
import datetime
from time import sleep
import openpyxl
import pyperclip
import pygetwindow
from tkinter import messagebox
import ddddocr
import pyautogui
from io import BytesIO
import subprocess
import sys
import json
import yaml
import threading
import shutil
import tkinter as tk

def load_config(config_path): # 读取yaml文件并返回一个字典
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            return config # 读出来是一个字典
    except FileNotFoundError:
        print(f'配置文件{config_path}不存在')
        return {}
    except yaml.YAMLError as e:
        print(f"YAML解析错误: {e}")
        return {}
config_constant = load_config('config_constant.yaml')

def save_json(dict, data_path): # 保存dict到json文件
    print(f'开始保存任务到{data_path}下的json文件')
    with open(data_path, 'w', encoding='utf-8') as f:
        json.dump(dict, f, ensure_ascii=False, indent=2, sort_keys=True)


def load_json(file_name): # 加载json文件得到dict字典
    if not os.path.exists(file_name): # 确保文件存在
        print(f'文件{file_name}不存在')
        return {} # 这里必须返回空字典而不是None，否则dict会变成None，对它使用字典相关方法的时候会报错

    try:
        with open(file_name, 'r', encoding='utf-8') as f:
            dict = json.load(f)
        print(f'加载{file_name}下json文件成功')
        return dict
    except Exception as e:
        print(f'json加载失败：{e}')
        return {} # 这里必须返回空字典而不是None，否则dict会变成None，对它使用字典相关方法的时候会报错


def get_current_path():
    """
    判断当前文件是打包为exe的还是源码，然后获取该文件所在的路径
    因为在不同的模式下需要不同的获取方式，否则可能出错
    :return:返回当前文件所在位置的路径
    """
    if getattr(sys, 'frozen', False):
        # 打包模式：使用 sys.executable 的目录
        print("检测到打包模式，使用 exe 所在目录")
        return os.path.dirname(sys.executable)
    else:
        # 源码模式：使用 __file__ 的目录
        print("检测到源码模式，使用脚本所在目录")
        return os.path.dirname(os.path.abspath(__file__))


def paste_write(string) -> None:
    """
    使用pyperclip这个库将要输入的文字存入剪切板中，然后粘贴。主要是因为使用typewrite是模拟键盘输入，如果输入法不对可能会影像输入结果
    :param string:想要输入的字符
    :return:无
    """
    pyperclip.copy(str(string))
    pyautogui.hotkey('ctrl','v',interval=0.1)


def find_img_center(img_path: str,confidence: float,max_loops: int,interval: float):
    """
    查找一个图片的中间坐标
    :param img_path: 图片路径
    :param confidence: 置信度
    :param max_loops: 最大循环次数
    :param interval: 每次查找间隔
    :return: 找到了返回位置信息；未找到或报错返回None
    """
    for i in range(max_loops):
        try:
            # 如果这里第一行捕获了异常，下面的return就会跳过，直接到对应的except去
            position = pyautogui.locateCenterOnScreen(img_path, confidence=confidence)  # 这里返回的是中心位置(x,y)
            return position
        except pyautogui.ImageNotFoundException:  # 如果没找到图片（正常情况）
            pass  # 空操作，继续执行，如果用continue下面的sleep就会被跳过了
        except Exception as e:  # 如果出现其它报错
            print(f'find_img出错：{e}')
            return None
        sleep(interval)  # 每次循环的间隔
    return None  # 如果一直都没找到图片，需要返回一个None


def find_img(img_path: str,confidence: float,max_loops: int,interval: float):
    """
    查找一个图片的左上角坐标与宽高
    :param img_path: 图片路径
    :param confidence: 置信度
    :param max_loops: 最大循环次数
    :param interval: 每次查找间隔
    :return: 找到了返回位置信息；未找到或报错返回None
    """
    for i in range(max_loops):
        try:
            # 如果这里第一行捕获了异常，下面的return就会跳过，直接到对应的except去
            position = pyautogui.locateOnScreen(img_path, confidence=confidence) # 这里返回的是元组(left,top,width,height)
            return position
        except pyautogui.ImageNotFoundException: # 如果没找到图片（正常情况）
            pass # 空操作，继续执行，如果用continue下面的sleep就会被跳过了
        except Exception as e: # 如果出现其它报错
            print(f'find_img出错：{e}')
            return None
        sleep(interval) # 每次循环的间隔
    return None # 如果一直都没找到图片，需要返回一个None


def mouse_move_center(current_img_path: str,img_name,max_time: int=0.2,confidence: float=0.9,interval: float=0.1):
    """
    让鼠标移动到识别图像中间的位置，同时需要指定置信度，并且结尾默认带了一个休眠1s
    :param current_img_path: 当前总图片文件夹路径
    :param img_name: 需要识别的图像名称
    :param confidence: 置信度
    :param max_time: 最大尝试时间
    :param interval: 每次尝试间隔
    :return: 找到了图片位置或未找到 和 寻找的图片路径
    """
    img_path = os.path.join(current_img_path, img_name)
    max_loops = max(2, int(max_time / interval))  # 确保至少尝试2次
    position = find_img_center(img_path, confidence, max_loops, interval)
    if position:
        pyautogui.moveTo(position)
        sleep(0.1)
    flag = bool(position)
    return flag,img_path


def mouse_move_not_center(current_img_path: str,img_name: str,dx: float,dy: float,max_time: int=0.2,confidence: float=0.9,interval: float=0.1):
    """
    让鼠标移动到识别图像左上角的位置并且移动给定的偏移量，同时需要指定置信度，并且结尾默认带了一个休眠1s
    :param current_img_path: 当前总图片文件夹路径
    :param img_name: 需要识别的图像名称
    :param dx: x偏移量
    :param dy: y偏移量
    :param confidence: 置信度
    :param max_time: 最大尝试时间
    :param interval: 每次尝试间隔
    :return: 找到了图片位置或未找到 和 寻找的图片路径
    """
    img_path = os.path.join(current_img_path, img_name)
    max_loops = max(2, int(max_time / interval))  # 确保至少尝试2次
    position_info = find_img(img_path, confidence, max_loops, interval)
    if position_info:
        x,y,_,_ = position_info
        position = (x + dx, y + dy)
        pyautogui.moveTo(position)
        sleep(0.1)
    flag = bool(position_info)
    return flag, img_path


def ocr(position: tuple) -> str: # OCR文字识别验证码
    x, y, w, h = position # 使用元组解包，x和y是图片左上角绝对坐标，w和h是图片的宽高
    screenshot = pyautogui.screenshot(region=(x, y, w, h)) # 使用pyautogui对指定地方截图，但不会保存下来，只是临时数据
    ocr = ddddocr.DdddOcr(show_ad=False) # 初始化ddddocr的识别引擎，show_ad是为了不显示他的广告
    img_bytes = BytesIO() # 创建一个内存中的二进制对象，用于临时存储图片数据
    screenshot.save(img_bytes, format='PNG') # 将screenshot（PIL.Image 对象）保存到内存中的 BytesIO流，png格式更适合验证码识别
    img_data = img_bytes.getvalue() # 从BytesIO流中提取二进制数据（bytes类型）
    result = ocr.classification(img_data) # 调用 ddddocr识别验证码图片
    return result


def try_ocr(current_img_path: str, error_img_name: str, ocr_position: tuple, captcha_position: tuple, login_button_name: str='',confidence: float=0.9, tries: int=10, interval: float=0.2):
    """
    用于尝试验证码并且直接登录
    :param current_img_path: 当前图片总路径
    :param error_img_name: 验证码识别错误图片
    :param ocr_position: 验证码实际位置，用于截图后ocr识别
    :param captcha_position: 验证码输入框位置
    :param login_button_name: 登录按钮图片名称，不写该参数就是回车登录，写了就是点击该按钮登录
    :param confidence: 验证码识别置信度
    :param tries: 最大尝试次数
    :param interval: 通用的操作之间间隔
    :return: 无
    """
    # 这里的captcha_position是为了拿到验证码输入的位置，方便重输
    error_img_path = os.path.join(current_img_path, error_img_name)
    for i in range(tries):
        captcha = ocr(ocr_position) # 获取验证码信息
        paste_write(captcha) # 输入验证码
        sleep(interval)
        if not login_button_name:
            pyautogui.press('enter') # 回车尝试登录
        else:
            mouse_move_center(current_img_path,login_button_name) # 移动到登录按钮处
            pyautogui.click()
        sleep(interval)
        # 判断是否有验证码报错
        flag = bool(find_img_center(error_img_path, confidence=confidence, max_loops=tries, interval=interval))
        if not flag: # 如果没找到那就说明成功登录，退出循环
            break
        sleep(0.5)
        pyautogui.doubleClick(*captcha_position)  # 选中验证码区域，这里*captcha_position是解包元组，把元组中俩值传给x和y
        sleep(0.1)
        pyautogui.hotkey('ctrl','a',interval=0.1) # 快捷键全选，确保一定全选了验证码，方便后面输入新的
        sleep(0.5)

def scale_to_150(): # 把网页调至150%大小
    pyautogui.hotkey('ctrl','0',interval=0.1)
    for i in range(3):
        pyautogui.hotkey('ctrl', '+', interval=0.1)
        sleep(0.05)

def get_time_standard(flag:str): # 获取时间标准格式，通过传入字符判断要今天/明天/昨天
    today = datetime.date.today()
    if flag == 'today':
        today_str = today.strftime("%Y-%m-%d")
        return today_str
    elif flag == 'yesterday':
        yesterday = today - datetime.timedelta(days=1)
        yesterday_str = yesterday.strftime("%Y-%m-%d")
        return yesterday_str
    elif flag == 'tomorrow':
        tomorrow = today + datetime.timedelta(days=1)
        tomorrow_str = tomorrow.strftime("%Y-%m-%d")
        return tomorrow_str

def print_file(path):
    os.startfile(path, 'print') # 直接打印指定文件

def open_path(path):
    os.startfile(path) # 直接打开指定文件或文件夹

def create_folder(path):
    try:
        os.mkdir(path) # 该方法用于创建单级文件夹
        print(f'文件夹{path}已创建成功')
    except FileExistsError:
        print(f'文件夹{path}已存在！')

def copy_paste_file(src, dst): # 复制粘贴文件并视情重命名，由于需要先创建文件夹再复制，所以记得先执行前面的创建文件夹操作
    try:
        shutil.copy2(src, dst)
        print(f'文件 {src} 已成功复制到 {dst}')
    except FileNotFoundError:
        print(f'源文件 {src} 不存在，请检查路径')
    except PermissionError:
        print('没有足够的权限进行复制操作')

def start_threading(func, task_name): # 使用多线程创建任务模板，同时里面提供弹窗信息显示当前任务
    # 显示弹窗信息
    info_win = tk.Toplevel()
    info_win.attributes('-topmost', True)  # 置顶
    info_win.attributes('-disabled', True)  # 禁用交互
    info_win.overrideredirect(True)  # 无边框
    # 窗口位置
    screen_width = info_win.winfo_screenwidth()
    screen_height = info_win.winfo_screenheight()
    x = 250  # 弹窗宽度
    y = 50  # 弹窗高度
    info_win.geometry(f'{x}x{y}+{screen_width - x - 10}+{screen_height - y - 40}')
    # 内容
    text = task_name + ' 任务正在执行\n请勿手动操作(弹窗除外)'
    tk.Label(info_win, text=text, font=("Arial", 10, "bold"), fg='red').pack(pady=5)

    def monitored_func(): # 监控任务，等任务结束，执行提示窗口也关闭
        try:
            func()
        finally:
            # 任务完成后（无论成功失败）关闭窗口
            if info_win.winfo_exists():
                info_win.after(0, info_win.destroy)

    # 实际线程执行
    thread = threading.Thread(target=monitored_func)
    thread.daemon = True  # 守护线程，主窗口关闭时自动结束
    thread.start()
