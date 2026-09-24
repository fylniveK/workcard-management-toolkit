import os.path
import tkinter as tk
from tkinter import messagebox, ttk
import ddddocr
import pyautogui
import openpyxl
import threading
from docx import Document
from docx.shared import Pt
from io import BytesIO

from general import *
import yaml
import json
import time
import datetime
import re
import copy
import webbrowser

from legacy_portal import employee
from others import *
from mro import *
from document_cloud import *
from tech_portal import *

root = tk.Tk()
root.title("智慧工卡管理平台")
root.geometry("1100x700+760+0")

# 统一设置微软雅黑字体和大小
style = ttk.Style()
style.configure(".", font=("Microsoft YaHei", 10))

# 创建主分割窗口，因为ttk中有这种专用功能，所以没有用tk
paned_window = ttk.PanedWindow(root, orient=tk.HORIZONTAL) # 专用于创建类似划分区域的效果
paned_window.pack(fill=tk.BOTH, expand=True)
# 左侧区域
left_frm = tk.Frame(paned_window)
paned_window.add(left_frm, weight=1)
# 右侧区域 - TODOlist
right_frm = tk.Frame(paned_window)
paned_window.add(right_frm, weight=2) # weight指的是两个frame宽度占比权重，左边分1份，右边2份

## 左侧再分为功能按钮区和信息区
left_paned = ttk.PanedWindow(left_frm, orient=tk.VERTICAL)
left_paned.pack(fill=tk.BOTH, expand=True)
function_button_frm = tk.Frame(left_paned) # 功能按钮区
left_paned.add(function_button_frm, weight=1)
info_frm = tk.Frame(left_paned) # 信息区
left_paned.add(info_frm, weight=1)

# todolist标题
todo_title = ttk.Label(
    right_frm,
    text="todolist",
    font=('Microsoft YaHei', 12, 'bold'),
    anchor='center'
)
todo_title.pack(fill=tk.X)
# 添加todolist功能按钮区域
todo_function_frm = ttk.Frame(right_frm)
todo_function_frm.pack(fill=tk.X, padx=2, pady=2)

## 右侧垂直分割为待办区和已完成区
right_paned = ttk.PanedWindow(right_frm, orient=tk.VERTICAL)
right_paned.pack(fill=tk.BOTH, expand=True)

# 配置PanedWindow显示分割线
style = ttk.Style()
style.configure("TPanedwindow", background='#BBB')  # 设置分割线颜色

def create_scrollable_frm(parent, title):
    '''
    用于创建一整个区域并加入滚动条
    container (主容器)
    ├── 标题标签
    └── canvas_frm (画布容器)
        ├── canvas (画布，视口)
        │   └── scrollable_frm (可滚动内容区，实际放任务的地方)
        └── scrollbar (滚动条)
    :param parent: 父级容器，新创建的滚动区域将放置在此容器内
    :param title: 区域标题文本，显示在滚动区域上方
    :return: 返回包含标题和滚动区域的完整容器对象，可通过 container.scrollable_frm 访问内部的可滚动Frame
    '''
    container = ttk.Frame(parent) # 创建一个frame作为整个滚动区域的外壳

    # 标题
    ttk.Label(container, text=title, font=('Microsoft YaHei', 10, 'bold')).pack(anchor='w', pady=(5, 10))

    # 滚动区域，必须得在canvas里才能创建滚动条
    canvas_frm = ttk.Frame(container) # 在container内创建一个新的Frame专门放Canvas和滚动条
    canvas_frm.pack(fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(canvas_frm) # 在canvas_frm中创建垂直滚动条
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y) # 放在右侧垂直方向填满

    # 创建画布，画布垂直滚动时调用滚动条的set方法
    canvas = tk.Canvas(canvas_frm, yscrollcommand=scrollbar.set, highlightthickness=0, bd=0)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=canvas.yview) # 滚动条被拖动时调用画布的yview方法

    scrollable_frm = ttk.Frame(canvas) # 创建实际放内容的frame
    # 在画布上创建一个“窗口”来放置这个frame，窗口西北角也就是左上角对齐画布0，0点
    canvas_window = canvas.create_window((0, 0), window=scrollable_frm, anchor='nw')

    # 绑定事件
    def on_frame_configure(event):
        # 当scrollable_frm大小改变时触发
        # canvas.bbox('all')：获取画布上所有内容的边界框
        # scrollregion设置画布可滚动范围
        canvas.configure(scrollregion=canvas.bbox('all'))

    def on_canvas_configure(event):
        # 当canvas大小改变时触发
        # event.width是画布的新宽度
        # itemconfig()是调整画布上窗口的宽度，让内容区宽度匹配画布
        canvas.itemconfig(canvas_window, width=event.width)

    # 当 scrollable_frm 大小变化时更新滚动区域
    scrollable_frm.bind('<Configure>', on_frame_configure)
    # 当 canvas 大小变化时调整内容区宽度
    canvas.bind('<Configure>', on_canvas_configure)

    container.scrollable_frm = scrollable_frm # 在 container 上添加属性，方便外部访问可滚动区域（使用的是引用存储技术标记了）
    container.canvas = canvas  # 将canvas也作为属性保存

    # 添加Windows鼠标滚轮支持
    def on_mousewheel(event):
        # Windows: 向上滚动为正值(120)，向下滚动为负值(-120)
        # 除以120是为了标准化，yview_scroll接收整数
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # 绑定鼠标滚轮事件到Canvas和可滚动区域
    canvas.bind("<MouseWheel>", on_mousewheel)
    scrollable_frm.bind("<MouseWheel>", on_mousewheel)

    return container

# 创建label+entry的结合框架
def create_label_and_entry(window, label_text, entry_text):
    frm = tk.Frame(window)
    frm.pack(anchor='w', fill='x', padx=1, pady=1, expand=False)

    tk.Label(frm, text=label_text, width=0).pack(side=tk.LEFT)
    entry = tk.Entry(frm)
    entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    entry.insert(0, entry_text)

    return frm, entry

### 上部分待办区划分（带滚动条）
pending_section = create_scrollable_frm(right_paned, "待办任务（回车创建任务）")
right_paned.add(pending_section ,weight=1)
pending_frm = pending_section.scrollable_frm # 拿到实际装任务的frame
### 下部分已完成任务划分（带滚动条）
completed_section = create_scrollable_frm(right_paned, "已完成任务")
right_paned.add(completed_section, weight=1)
completed_frm = completed_section.scrollable_frm # 拿到实际装任务的frame

# 左侧功能按钮区域
## 创建每个区域的内容框架，后续需要添加区域就来这里，这里用字典把他们打包是方便后续同一处理
area_frms = {}
# legacy_portal_frm = ttk.Frame(function_button_frm) # 旧版维修平台功能区 已停用
# AMMS_frm = ttk.Frame(function_button_frm) # AMMS功能区 已停用
cloud_frm = ttk.Frame(function_button_frm) # 文档云功能区
MRO_frm = ttk.Frame(function_button_frm) # MRO功能区
TECH_PORTAL_frm = ttk.Frame(function_button_frm) # MRO功能区
quick_frm = ttk.Frame(function_button_frm) # 快捷操作区
todo_frm = ttk.Frame(function_button_frm) # todolist区
guide_frm = ttk.Frame(function_button_frm) # 使用指南区
# area_frms['LEGACY_PORTAL'] = legacy_portal_frm # 已停用
# area_frms['AMMS'] = AMMS_frm # 已停用
area_frms['cloud'] = cloud_frm
area_frms['MRO'] = MRO_frm
area_frms['TECH_PORTAL'] = TECH_PORTAL_frm
area_frms['quick'] = quick_frm
area_frms['todo'] = todo_frm
area_frms['guide'] = guide_frm

def show_area(area_name:str, area_frms:dict=area_frms):
    '''
    根据选择按钮不同显示不同功能区域
    :param area_name:功能区名字
    :param area_frms:包含所有功能区的字典
    :return:无
    '''
    # 隐藏所有区域
    for frm in area_frms.values():
        frm.pack_forget()

    # 显示选中的区域
    area_frms[area_name].pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

## 创建上面切换按钮区域，它必须要在show_area默认的上面，因为pack打包是有顺序的，如果反了就会导致切换按钮在最下面显示而不是最上面
switch_frm = ttk.Frame(function_button_frm)
switch_frm.pack(fill=tk.X, padx=5, pady=5)

## 默认显示待办模板区域，它必须要在切换按钮的下面
show_area('todo', area_frms)

## 创建每个区域对应的切换按钮，同时把函数绑定进来，让他们每次只显示自己的区域
# ttk.Button(switch_frm, text='旧版维修平台', width=0, command=lambda: show_area('LEGACY_PORTAL')).pack(side=tk.LEFT) #已停用
# ttk.Button(switch_frm, text='AMMS', width=0, command=lambda: show_area('AMMS')).pack(side=tk.LEFT) #已停用
ttk.Button(switch_frm, text='文档云', width=0, command=lambda: show_area('cloud')).pack(side=tk.LEFT)
ttk.Button(switch_frm, text='MRO', width=0, command=lambda: show_area('MRO')).pack(side=tk.LEFT)
ttk.Button(switch_frm, text='TECH_PORTAL', width=0, command=lambda: show_area('TECH_PORTAL')).pack(side=tk.LEFT)
ttk.Button(switch_frm, text='快捷小工具', width=0, command=lambda: show_area('quick')).pack(side=tk.LEFT)
ttk.Button(switch_frm, text='待办模板', width=0, command=lambda: show_area('todo')).pack(side=tk.LEFT)
ttk.Button(switch_frm, text='使用指南', width=0, command=lambda: show_area('guide')).pack(side=tk.LEFT)





# 左侧下面信息区info_frm***
## 标题
info_title = tk.Label(
    info_frm,
    text="***今日A检信息（请先核对）***",
    font=('Microsoft YaHei', 12, 'bold'),
    anchor='center',
    fg = 'red'
)
info_title.pack(fill=tk.X, padx=5, pady=1)

## 实际显示区
details_frm = tk.Frame(info_frm)
details_frm.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

# root窗口定检信息的功能按钮区
## details_dict字典用于：制定定检信息的默认配置&临时存储tk的组件frm和entry，不存储info
details_dict = {
    'AIRCRAFT_ID': {'name': '飞机号：', 'prompt': '机尾号（例如B-1234可以只填1234）：'},
    'CHECK_NO': {'name': 'A检级别：', 'prompt': 'A检级别（例如A120可以只填120，或填“换发”）：'},
    'EMPLOYEE_NO': {'name': '领料员工号：', 'prompt': '领料员工号（填自己即可）：'},
    'DIRECTOR_NAME': {'name': '值班主任姓名：', 'prompt': '值班主任姓名：'},
    'FLIGHT_NO': {'name': '航班号：', 'prompt': '航班号（选填，例如1234）：'},
    'BRIDGE_NO': {'name': '桥位：', 'prompt': '桥位：'},
    'GUN': {'name': '有效号：', 'prompt': '有效号：'},
    'FLIGHT_LOG_PAGE': {'name': '飞行记录本页号(有保留就填)：', 'prompt': '飞行记录本页号（有保留就填）：'},
    'ORDER_NO': {'name': '航材指令号：', 'prompt': '航材指令号：'}
}
# 创建一个用于json存储的新字典details_json_dict，专用于临时存储info并保存至json
# 主要是因为details_dict中包括了frm和entry这两个tk专属的对象，它们无法被序列化，json无法存储
# 这里必须先有个默认版本，否则如果在创建GUI时它是空的，就会导致完全没有输入框，死循环，而且要用深复制，副本不会影响原对象
details_json_dict = copy.deepcopy(details_dict) # 如果每次都在保存定检信息后使用这些信息，那么应该就可以从details_json_dict中获取信息而不是details_dict
# 指定定检信息的json存储地址
check_info_path = os.path.join(current_script_dir, 'data', 'check_info.json')


## 建立A检信息区中的按钮功能区
details_funciton_frm = tk.Frame(details_frm)
details_funciton_frm.pack(fill=tk.X, padx=1, pady=10)
# 创建定检信息label展示区
show_check_info_frm = tk.Frame(details_frm)
show_check_info_frm.pack(fill=tk.BOTH, padx=1, pady=1, expand=True)

def show_check_info():
    global details_json_dict
    # 先清空现有的所有label
    for widget in show_check_info_frm.winfo_children():
        widget.destroy()
    # 读取json中信息到details_json_dict中
    load_details_data = load_json(check_info_path)
    details_json_dict.update(load_details_data) # 使用更新的方法把数据合并覆盖到details_json_dict中，而不是直接用读取的数据取代details_json_dict，读取的数据可能是空
    # 创建A检信息区中的所有展示label，不可直接修改
    for key,value in details_json_dict.items():
        tk.Label(show_check_info_frm, text=f'{value.get("name")}{value.get("info")}', font=('Microsoft YaHei', 12), anchor='w').pack(anchor='w', fill='x', padx=1, pady=1, expand=False)

def create_edit_main_area(window):
    # 创建edit_check_info_window的具体编辑输入界面
    global details_dict

    for key, value in details_json_dict.items():
        label_text = value.get('prompt','读取出错，请检查') # 这里最好给一个默认值为空字符串，否则找不到返回None的话，下面的判断会报错，None无法去空格等操作
        entry_text = value.get('info','')

        frm, entry = create_label_and_entry(window, label_text, entry_text)
        details_dict[key]['frm'] = frm
        details_dict[key]['entry'] = entry


def clear_check_info(): # 遍历字典中所有entry，并删除其中的内容
    if messagebox.askyesno('清空所有信息', '清空信息无法还原，确认清空？'):
        for value in details_dict.values():
            entry = value.get('entry', None)
            if entry:
                entry.delete(0, tk.END) # 删除从开始到结束的所有内容
            else:
                print('在details_dict中未找到某个entry')


def save_check_info():
    global details_dict, details_json_dict

    for key, value in details_dict.items():
        entry = value.get('entry', None) # 获取entry
        if entry:
            entry_text = entry.get() # 提取信息文字
        else:
            print('获取entry出错')
            break
        details_json_dict[key]['info'] = entry_text # 把提取到的信息保存到字典的info中
        label_text = value.get('prompt','读取出错，请检查')

        ## 把飞机号、A检级别等转换为标准格式
        if label_text == '机尾号（例如B-1234可以只填1234）：': # 如果是飞机号这一条
            entry_text = entry_text.strip().upper() # 去除空格并且全转换为大写
            entry_text = re.sub(r'[^A-Z0-9]', '', entry_text) # 清除除了大写字母和数字外的内容
            pattern = r'^B?([A-Z0-9]+)$' # 匹配模式，拿出后面的机尾号
            match = re.match(pattern, entry_text) # 取得匹配值
            if match:
                body = match.group(1)
                entry_text = f'B-{body}'
            else: entry_text = ''

        elif label_text == 'A检级别（例如A120可以只填120，或填“换发”）：': # 如果是A检级别这一条
            entry_text = entry_text.strip() # 去除空格
            if '换发' in entry_text: # 此时为换发的情况
                entry_text = '换发'
            else: # 此时为A检的情况
                entry_text = entry_text.upper() # 全转换为大写
                entry_text = re.sub(r'[^A-Z0-9]', '', entry_text) # 清除除了大写字母和数字外的内容
                pattern = r'^A?(\d+)$' # 匹配模式，拿出后面的A检号
                match = re.match(pattern, entry_text) # 取得匹配值
                if match:
                    body = match.group(1)
                    entry_text = f'A{body}'
                else: entry_text = ''

        elif label_text == '航班号（选填，例如1234）：': # 如果是航班号这一条
            entry_text = entry_text.strip().upper()  # 去除空格并转为大写
            # 提取所有数字
            digits = re.sub(r'\D', '', entry_text)  # 移除非数字字符
            if digits:  # 如果有数字
                entry_text = f"{config_constant['constant']['flight_prefix']}{digits}"
            else:  # 如果没有数字
                entry_text = ''

        ## 由于details_dict中包括了frm和entry这两个tk专属的对象，它们无法被序列化，json无法存储
        ## 所有我们把之前字典的键以及值中的name、prompt、info信息存入新的字典details_json_dict，最后再导入json
        details_json_dict[key] = {
            'name': value['name'],
            'prompt': value['prompt'],
            'info': entry_text
        }

    ## 将details_json_dict存入json文件
    save_json(details_json_dict, check_info_path)
    show_check_info()


def edit_check_info(): ## 修改定检信息
    edit_check_info_window = tk.Toplevel(root)
    edit_check_info_window.title('修改定检信息')
    edit_check_info_window.geometry('450x350+960+350')

    # 把edit_check_info_window作为root的子窗口，这样就可以正常聚焦置顶，同时主窗口最小化子窗口也会一起最小化
    edit_check_info_window.transient(root)  # 子窗口

    tk.Label(
        edit_check_info_window,
        text="***修改飞机信息后,关闭该小窗口自动保存***",
        font=('Microsoft YaHei', 12, 'bold'),
        anchor='center',
        fg='blue'
    ).pack(fill=tk.X, padx=5, pady=1)

    ## edit_check_info_window窗口的功能按钮区
    edit_check_info_function_frm = tk.Frame(edit_check_info_window)
    edit_check_info_function_frm.pack(fill=tk.X, padx=1, pady=5)
    ttk.Button(edit_check_info_function_frm, text='清空所有信息', command=clear_check_info, width=0).pack(anchor='n', side=tk.LEFT, padx=1)

    ## 创建edit_check_info_window的具体编辑输入界面
    create_edit_main_area(edit_check_info_window)

    def close_edit_check_info():
        save_check_info()
        edit_check_info_window.destroy()
    ## 关闭edit_check_info_window时绑定保存信息函数
    edit_check_info_window.protocol('WM_DELETE_WINDOW', close_edit_check_info)

def copy_check_info(flag):
    if flag == 'order_no':
        order_no = details_json_dict.get('ORDER_NO').get('info', None) # 获取指令号
        if order_no:
            pyperclip.copy(str(order_no))
        else:
            messagebox.showwarning('注意','输入指令号后再使用该功能')
    elif flag == 'flight_log_page':
        flight_log_page = details_json_dict.get('FLIGHT_LOG_PAGE').get('info', None) # 获取飞行记录本页号
        if flight_log_page:
            pyperclip.copy(str(flight_log_page))
        else:
            messagebox.showwarning('注意','输入飞行记录本页号后再使用该功能')
    else:
        print('输入的参数不符合copy_check_info函数标准')

## 绑定“修改信息”按钮与功能函数
ttk.Button(details_funciton_frm, text='修改信息', command=edit_check_info, width=0).pack(anchor='n', side=tk.LEFT, padx=1)
ttk.Button(details_funciton_frm, text='复制指令号', command=lambda:copy_check_info('order_no'), width=0).pack(anchor='n', side=tk.LEFT, padx=1)
ttk.Button(details_funciton_frm, text='复制飞行记录本页号', command=lambda:copy_check_info('flight_log_page'), width=0).pack(anchor='n', side=tk.LEFT, padx=1)








# 右侧todolist区域
## 生成唯一 任务ID
def generate_task_id():
    '''基于时间戳生成唯一任务ID'''
    timestamp = int(time.time() * 1000)
    return timestamp

def general_create_task(state=False, text='', id=''): # 可复用的创建任务方法
    if state: # 如果任务是完成状态
        # 创建框架
        task_frm = ttk.Frame(completed_frm)
        task_frm.pack(fill=tk.X, padx=2, pady=2)

        # 确定对应的Canvas
        target_canvas = completed_section.canvas

        # 切换按钮（此时为撤回按钮）
        toggle_btn = ttk.Button(task_frm, text='撤回', command=lambda task_frm=task_frm: toggle(task_frm)) #这里使用匿名函数把框架传出去，必须写形参否则会出问题
        toggle_btn.pack(side=tk.LEFT)

    else:
        task_frm = ttk.Frame(pending_frm)
        task_frm.pack(fill=tk.X, padx=2, pady=2)

        # 确定对应的Canvas
        target_canvas = pending_section.canvas

        # 切换按钮（此时为完成按钮）
        toggle_btn = ttk.Button(task_frm, text='完成',command=lambda task_frm=task_frm: toggle(task_frm))  # 这里使用匿名函数把框架传出去，必须写形参否则会出问题
        toggle_btn.pack(side=tk.LEFT)

    # 输入框
    entry = ttk.Entry(task_frm)
    entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    entry.insert(0, text)

    # 删除按钮
    delete_btn = ttk.Button(task_frm, text='删除',command=lambda task_frm=task_frm: delete(task_frm))  # 这里使用匿名函数把框架传出去，必须写形参否则会出问题
    delete_btn.pack(side=tk.RIGHT)

    # 生成唯一id
    if not id:
        id = generate_task_id()
    # 在frame上存储数据引用（标记），这一步非常关键
    # 这里并不是创建了一个task_data变量，而是在对象上动态添加了属性
    task_frm.id = id
    task_frm.state = state
    task_frm.toggle = toggle_btn
    task_frm.entry = entry
    task_frm.delete = delete_btn

    # 为任务框架及其所有子控件绑定鼠标滚轮
    def on_task_mousewheel(event):
        # 使用上面确定的Canvas
        target_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"  # 阻止事件继续传播

    # 为任务框架本身绑定
    task_frm.bind("<MouseWheel>", on_task_mousewheel)

    # 为任务框架内的所有子控件绑定
    for child in task_frm.winfo_children():
        child.bind("<MouseWheel>", on_task_mousewheel)

    return task_frm

def add_new_task(text=''): # 创建一个全新的任务，和回车绑定
    task_frm = general_create_task()
    task_frm.entry.focus() # 让光标聚焦
    pending_section.canvas.update_idletasks()  # 更新UI
    pending_section.canvas.yview_moveto(1.0)  # 滚动到底部（1.0表示100%位置），这样创建新任务之后就会自动滚动下去
root.bind('<Return>', lambda e:add_new_task()) # 回车绑定新建任务

def toggle(task_frm): # 任务状态改变->切换任务的区域
    id = task_frm.id
    state = not task_frm.state # 直接切换状态
    text = task_frm.entry.get()
    task_frm.destroy() # 删除之前的任务框架

    task_frm = general_create_task(state, text, id) # 用原有信息创建新的任务框架

def delete(task_frm):
    if messagebox.askyesno('删除该任务', '是否确认删除该任务，操作不可恢复'):
        task_frm.destroy()

tasks_dict = {} # 创建一个空字典后续存储任务信息
def collect_all_tasks_info(tasks_dict): # 把所有任务信息收集到tasks_dict字典中
    tasks_dict.clear() # 先清空所有旧信息
    # 遍历所有任务框架
    all_task_frames = []
    all_task_frames.extend(pending_frm.winfo_children()) # 待办区
    all_task_frames.extend(completed_frm.winfo_children()) # 完成区

    # 获取每个框架中任务信息并收集到任务字典中
    for task_frm in all_task_frames:
        id = task_frm.id
        state = task_frm.state
        text = task_frm.entry.get()
        tasks_dict[id] = {
            'state': state,
            'text': text
        }
    return tasks_dict

## 使用json记录&提取todolist数据，存在todo_data.json
todo_data_path = os.path.join(current_script_dir, 'data', 'todo_data.json')

def close_save(tasks_dict): # 关闭时自动保存任务信息
    tasks_dict = collect_all_tasks_info(tasks_dict)
    save_json(tasks_dict, todo_data_path)
    root.destroy()

def restore_tasks(tasks_dict): # 从保存的json文件中复现所有任务
    if not tasks_dict:
        print('无json文件 或 json为空 或 json读取出错')
    else:
        for task_id, task_info in tasks_dict.items():
            # 生成的时间戳用的是int类型，但是由于时间戳是键，保存在json后键会被强制转换为str类型，重新打开GUI从json中读取信息时，json的键是字符串
            # 此时如果再新建任务，新任务的时间戳是int类型，但是由于我们进行了任务排序，时间戳这里会进行比较，int和str无法比较
            # 故下面需要先把json中读取的键全部从str转换为int，这样才能比较，否则报错
            if isinstance(task_id, str) and task_id.isdigit(): # 拿到的id（时间戳）判断是否是str类型 并且 全由数字组成
                task_id = int(task_id) # 把它们全转换为int
            state = task_info['state']
            text = task_info['text']
            task_frm = general_create_task(state, text, task_id)

def clear_all_tasks(): # 清除所有任务信息，但不删除json中信息
    if messagebox.askyesno('删除所有任务', '是否确认删除所有任务，操作不可恢复'):
        # 遍历所有任务框架
        all_task_frames = []
        all_task_frames.extend(pending_frm.winfo_children())  # 待办区
        all_task_frames.extend(completed_frm.winfo_children())  # 完成区

        # 获取每个框架中任务信息并摧毁
        for task_frm in all_task_frames:
            task_frm.destroy()

def manual_save(tasks_dict): # 手动保存任务信息
    tasks_dict = collect_all_tasks_info(tasks_dict)
    # print(tasks_dict) # 查看tasks_dict的内容，是否正常
    save_json(tasks_dict, todo_data_path)

# 添加todolist实际功能按钮
ttk.Button(todo_function_frm, text='保存', width=0, command=lambda: manual_save(tasks_dict)).pack(side=tk.LEFT, padx=1)
ttk.Button(todo_function_frm, text='清除所有任务', width=0, command=clear_all_tasks).pack(side=tk.LEFT, padx=1)







# 左上角快捷功能部分
# cloud功能区域：文档云登录（作为其它登录的基础）&录工时
def cloud_login():
    mro_user_name, mro_password, mro_mel_license = mro_account()
    document_cloud_auto_login(mro_user_name, mro_password)

def record_man_hours():
    os.startfile(auto_man_hours_template_path)
    os.startfile(auto_man_hours_fetch_path)
    webbrowser.open(man_hours_collect_url)
    sleep(2)
    # 找到所有相关浏览器窗口（标题包含关键字即可）并最大化
    windows = gw.getWindowsWithTitle('工时收集表')
    if windows:
        browser = windows[0]  # 打开其中第一个，好像就是打开的最新的那一个
        browser.maximize()  # 最大化
        browser.activate()  # 聚焦




# MRO功能区域：MRO登录&MRO中各种操作
def mro_login(): # MRO登录
    if messagebox.askyesno('是否已经使用“登录文档云”功能?','若已使用，点击【是】登录MRO\n若未使用，请点击【否】然后先执行“登录文档云”功能'):
        mro_user_name, mro_password, mro_mel_license = mro_account()
        mro_auto_login(mro_user_name, mro_password)

def query_aircraft_parts(): # 航材库存查询
    if messagebox.askyesno('是否已使用“登录MRO”功能','若已使用，点击【是】打开查询航材网址\n若未使用，请点击【否】后先执行“登录MRO”功能'):
        mro_auto_query_aircraft_parts()

def mro_tasks_issue():
    aircraft_id = details_json_dict.get('AIRCRAFT_ID').get('info', None)  # 获取飞机号
    if not aircraft_id:
        messagebox.showinfo('注意','请先输入飞机号')
        return
    if messagebox.askyesno('是否已使用“登录MRO”功能','若已使用，点击【是】打开任务派发网址\n若未使用，请点击【否】后先执行“登录MRO”功能'):
        mro_auto_tasks_issue(aircraft_id)

def mro_non_routine_workcard_search():
    aircraft_id = details_json_dict.get('AIRCRAFT_ID').get('info', None)  # 获取飞机号
    if not aircraft_id:
        messagebox.showinfo('注意', '请先输入飞机号')
        return
    if messagebox.askyesno('是否已使用“登录MRO”功能', '若已使用，点击【是】打开非卡方案网址\n若未使用，请点击【否】后先执行“登录MRO”功能'):
        mro_auto_non_routine_workcard_search(aircraft_id)



# TECH_PORTAL功能区域：TECH_PORTAL登录&TECH_PORTAL中各种操作
def tech_portal_login():
    mro_user_name, mro_password, mro_mel_license = mro_account()
    tech_portal_auto_login(mro_user_name, mro_password)

def tech_portal_737ng_pp_open():
    if messagebox.askyesno('是否已使用“登录TECH_PORTAL”功能', '若已使用，点击【是】打开737NG-PP手册\n若未使用，请点击【否】后先执行“登录TECH_PORTAL”功能'):
        tech_portal_737ng_pp_auto_open()

def tech_portal_737ng_ipc_open():
    if messagebox.askyesno('是否已使用“登录TECH_PORTAL”功能', '若已使用，点击【是】打开737NG-IPC手册\n若未使用，请点击【否】后先执行“登录TECH_PORTAL”功能'):
        tech_portal_737ng_ipc_auto_open()



# quick功能区域：一键打开&创建等操作
## 一键创建今日文件夹&航材需求单
def create_check_dir_excel():
    aircraft_id = details_json_dict.get('AIRCRAFT_ID').get('info', None) # 获取飞机号
    check_no = details_json_dict.get('CHECK_NO').get('info', None) # 获取定检级别
    director_name = details_json_dict.get('DIRECTOR_NAME').get('info', None) # 获取主任姓名
    order_no = details_json_dict.get('ORDER_NO').get('info', None) # 获取指令号
    employee_no = details_json_dict.get('EMPLOYEE_NO').get('info', None) # 获取领料员工号
    bridge_no = details_json_dict.get('BRIDGE_NO').get('info', None) # 获取桥位

    if aircraft_id and check_no and director_name and order_no and employee_no and bridge_no: # 如果有飞机号/A检级别主任姓名/指令号/领料员工号/桥位
        name = f'{aircraft_id}_{check_no}'  # 根据今天飞机号和A检级别命名
        check_dir_path = os.path.join(desktop_path, name)
        if not os.path.exists(check_dir_path):  # 如果没有创建文件夹
            create_folder(check_dir_path)  # 创建文件夹
        HangCaiXuQiuDan_dst = os.path.join(check_dir_path, f'{name}航材需求单.xlsx')  # 新建航材需求单的路径
        if not os.path.exists(HangCaiXuQiuDan_dst):  # 如果不存在航材需求单
            copy_paste_file(HangCaiXuQiuDan_template_path, HangCaiXuQiuDan_dst)  # 复制航材需求单到指定位置
            today = datetime.datetime.now()  # 获取今天日期

            wb = openpyxl.load_workbook(HangCaiXuQiuDan_dst)  # 加载Excel文件到内存
            ws = wb.active  # 获取活动工作表(无需真的打开excel)
            current_text = ws['B2'].value  # 获取B2格（表头）内容
            if '月日' in current_text:
                current_text = current_text.replace('月日', f'{today.month}月{today.day}日')
            if "当日车间值班领导：" in current_text:
                current_text += director_name
            ws['B2'].value = current_text

            # 填写下面航材的基础信息
            ws['B4'].value = '1'
            ws['C4'].value = order_no
            ws['D4'].value = aircraft_id
            ws['H4'].value = employee_no
            ws['I4'].value = bridge_no
            ws['J4'].value = 'AOG'
            # 保存到新文件（或覆盖原文件）
            wb.save(HangCaiXuQiuDan_dst)
        else:
            messagebox.showwarning('注意', f'{name}航材需求单.xlsx已存在，确认删除后再使用本功能')
        check_txt_path = os.path.join(check_dir_path, f'{name}.txt')  # 新建txt的路径
        if not os.path.exists(check_txt_path):  # 如果不存在txt就创建一个
            try:
                with open(check_txt_path, 'w', encoding='utf-8') as f:  # 使用 'w' 模式，如果文件存在会覆盖
                    f.write(f'{name}\n航材指令号：{order_no}\n\n')  # 写入今天飞机号 & A检级别 & 航材指令号
                print(f"文件 {check_txt_path} 创建/写入成功。")
            except IOError as e:
                print(f"操作文件时出错: {e}")

    else: # 如果没有飞机号和A检级别
        messagebox.showwarning('注意','未填写完 飞机号/A检级别/值班主任姓名/指令号/领料员工号/桥位 信息，全部填写后再使用本功能')

## 一键创建今日文件夹&工作交接单，这里都没修改，需要修改
def create_check_dir_word():
    aircraft_id = details_json_dict.get('AIRCRAFT_ID').get('info', None) # 获取飞机号
    check_no = details_json_dict.get('CHECK_NO').get('info', None) # 获取定检级别
    director_name = details_json_dict.get('DIRECTOR_NAME').get('info', None) # 获取主任姓名
    bridge_no = details_json_dict.get('BRIDGE_NO').get('info', None) # 获取桥位

    if aircraft_id and check_no and director_name and bridge_no: # 如果有飞机号/A检级别/主任姓名/桥位信息
        name = f'{aircraft_id}_{check_no}' # 根据今天飞机号和A检级别命名
        check_dir_path = os.path.join(desktop_path, name)
        if not os.path.exists(check_dir_path): # 如果没有创建文件夹
            create_folder(check_dir_path) # 创建文件夹
        JiaoJieDan_dst = os.path.join(check_dir_path, f'{name}定检工作交接单.docx') # 新建工作交接单的路径
        if not os.path.exists(JiaoJieDan_dst): # 如果不存在工作交接单
            copy_paste_file(JiaoJieDan_template_path, JiaoJieDan_dst) # 复制工作交接单到指定位置
            today = datetime.datetime.now() # 获取今天日期
            date_str = f'{today.year}年{today.month}月{today.day}日'

            # 往docx里面写东西
            doc = Document(JiaoJieDan_dst)
            for paragraph in doc.paragraphs:
                if '年月日' in paragraph.text:
                    paragraph.text = paragraph.text.replace('年月日', date_str)

            table = doc.tables[0] # 获取文档内第一个表格
            # 因为表格有单元格合并，所以实际位置并非看上去那样，以我写的为准
            table.cell(0,1).text = aircraft_id # 第一行第二列填入飞机号
            table.cell(0,4).text = check_no # 第一行第五列填入A检级别
            table.cell(0,6).text = bridge_no # 第一行第七列填入桥位
            table.cell(1,1).text = director_name # 第二行第二列填入主任姓名
            # 保存（覆盖）交接单
            doc.save(JiaoJieDan_dst)

        else: # 已存在工作交接单
            messagebox.showwarning('注意', f'{name}定检工作交接单.docx已存在，确认删除后再使用本功能')
    else: # 如果没有飞机号/A检级别/主任姓名/桥位信息
        messagebox.showwarning('注意','未填写完 飞机号/A检级别/主任姓名/桥位 信息，全部填写后再使用本功能')

# 打印宵夜单
def print_XiaoYeDan():
    XYD_window = tk.Toplevel(root)
    XYD_window.title('打印宵夜单')
    XYD_window.geometry("300x50+900+209")

    def print_40():
        print_file(XiaoYeDan_file_path_40)
    def print_50():
        print_file(XiaoYeDan_file_path_50)
    def print_60():
        print_file(XiaoYeDan_file_path_60)

    button_frame = tk.Frame(XYD_window)
    button_frame.pack(expand=True, fill=tk.BOTH)
    print_40_button = tk.Button(button_frame, text='40人份', command=print_40)
    print_40_button.pack(side=tk.LEFT,padx=40)
    print_50_button = tk.Button(button_frame, text='50人份', command=print_50)
    print_50_button.pack(side=tk.LEFT,padx=0)
    print_60_button = tk.Button(button_frame, text='60人份', command=print_60)
    print_60_button.pack(side=tk.LEFT,padx=40)

# 一键搞定外网登录
def auto_login_online():
    mro_user_name, mro_password, mro_mel_license = mro_account()
    login_online(mro_user_name, mro_password)




# todolist功能区域：从模板创建任务
## 一键创建固定任务
def create_constant_tasks():
    todo_list = config_constant.get("todo_list")
    for i in range(0,len(todo_list)):
        text = todo_list[i]
        general_create_task(text=text)
## 创建非固定任务，这里大部分都需要借还的
tools_position = config_constant.get('tools_position')
def create_simple_task(text_list, tools_position=tools_position): # 创建单个简单任务，只拿不用还的
    list = []
    for key in text_list:
        t = tools_position.get(key)
        list.append(t)
    general_create_task(text=' & '.join(list))

def create_borrow_and_return_tasks(main_list:list, borrow_list:list=[], return_list:list=[], tools_position=tools_position): # 创建借还同步创建的任务
    # 提取固定部分，借与还都要用
    constant_text_list = []
    for key1 in main_list:
        t1 = tools_position.get(key1)
        constant_text_list.append(t1) # 收集每个字符串到列表中
    constant_text = ' & '.join(constant_text_list) # 使用join方法把列表中每个元素用 & 连接起来

    # 借的部分
    borrow_text = '借：' + constant_text
    if borrow_list:
        borrow_text_list = []
        for key2 in borrow_list:
            t2 = tools_position.get(key2)
            borrow_text_list.append(t2)
        borrow_text = borrow_text + '。' + '并且要' + ' & '.join(borrow_text_list)

    # 还的部分
    return_text = '还：' + constant_text
    if return_list:
        return_text_list = []
        for key3 in return_list:
            t3 = tools_position.get(key3)
            return_text_list.append(t3)
        return_text = return_text + '。' + '并且要' + ' & '.join(return_text_list)

    # 创建借&还的任务
    general_create_task(text=borrow_text)
    general_create_task(text=return_text)


# guide功能区域：各种使用指南，但也许这里都不需要写函数，基本都是打开网页



# 添加快捷按钮区域实际功能按钮
## 旧版维修平台中的内容按钮，已停用
# ttk.Button(legacy_portal_frm, text='旧版维修平台已停用', width=0).pack(anchor='n', side=tk.LEFT, padx=5, pady=2)
## AMMS中的内容按钮，已停用
# ttk.Button(AMMS_frm, text='AMMS已停用', width=0).pack(anchor='n', side=tk.LEFT, padx=5, pady=2)

## cloud中的内容按钮
ttk.Button(cloud_frm, text='登录文档云', width=0, command=lambda:start_threading(cloud_login,'登录文档云')).grid(row=0, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(cloud_frm, text='工时自动化', width=0, command=record_man_hours).grid(row=0, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(cloud_frm, text='质量整改', width=0, command=lambda:webbrowser.open(quality_rectification_url)).grid(row=0, column=2, padx=5, pady=2, sticky='nsew')

## MRO中的内容按钮
ttk.Button(MRO_frm, text='登录MRO', width=0, command=lambda:start_threading(mro_login,'登录MRO')).grid(row=0, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(MRO_frm, text='航材库存查询', width=0, command=query_aircraft_parts).grid(row=0, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(MRO_frm, text='任务派发', width=0, command=lambda:start_threading(mro_tasks_issue,'任务派发')).grid(row=0, column=2, padx=5, pady=2, sticky='nsew')
ttk.Button(MRO_frm, text='非卡方案', width=0, command=lambda:start_threading(mro_non_routine_workcard_search,'非卡方案')).grid(row=0, column=3, padx=5, pady=2, sticky='nsew')

## TECH_PORTAL中的内容按钮
ttk.Button(TECH_PORTAL_frm, text='登录TECH_PORTAL', width=0, command=lambda:start_threading(tech_portal_login,'登录TECH_PORTAL')).grid(row=0, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(TECH_PORTAL_frm, text='PP手册', width=0, command=lambda:start_threading(tech_portal_737ng_pp_open,'自动打开PP手册')).grid(row=0, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(TECH_PORTAL_frm, text='IPC手册', width=0, command=lambda:start_threading(tech_portal_737ng_ipc_open,'自动打开IPC手册')).grid(row=0, column=2, padx=5, pady=2, sticky='nsew')


## quick中的内容按钮
ttk.Button(quick_frm, text='创建文件夹&航材', width=0, command=create_check_dir_excel).grid(row=0, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='创建交接单', width=0, command=create_check_dir_word).grid(row=0, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='CFM、LEAP试车报告单', width=0, command=lambda: open_path(ng_ShiCheDan_file_path)).grid(row=1, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='常用件号', width=0, command=lambda: open_path(ChangYongJianHao_file_path)).grid(row=1, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='打宵夜单', width=0, command=print_XiaoYeDan).grid(row=1, column=2, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='本地IPC', width=0, command=lambda: open_path(Local_IPC_dir_path)).grid(row=1, column=3, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='909试车报告单', width=0, command=lambda: open_path(C909_ShiCheDan_file_path)).grid(row=2, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='909件号', width=0, command=lambda: open_path(C909_ChangYongJianHao_file_path)).grid(row=2, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='上网登录', width=0, command=lambda: start_threading(auto_login_online,'上网登录')).grid(row=2, column=2, padx=5, pady=2, sticky='nsew')
ttk.Button(quick_frm, text='客舱标牌件号', width=0, command=lambda: open_path(KeCangBiaoPai_file_path)).grid(row=2, column=3, padx=5, pady=2, sticky='nsew')

## todolist中的内容按钮，使用grid布局，以免超出边界，目前是最多四列的排布
ttk.Button(todo_frm, text='创建固定任务', width=0, command=create_constant_tasks).grid(row=0, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='IDG加油器', width=0, command=lambda: create_borrow_and_return_tasks(['IDG加油器'])).grid(row=0, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='周卡', width=0, command=lambda: create_borrow_and_return_tasks(['放水车','放油杆&放油瓶'])).grid(row=0, column=2, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='换前轮', width=0, command=lambda: create_borrow_and_return_tasks(['前轮'], return_list=['给挂签'])).grid(row=0, column=3, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='换主轮', width=0, command=lambda: create_borrow_and_return_tasks(['主轮'], return_list=['给挂签'])).grid(row=1, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='换轮刹', width=0, command=lambda: create_borrow_and_return_tasks(['轮刹'], borrow_list=['轮刹封圈'], return_list=['给挂签'])).grid(row=1, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='试大车', width=0, command=lambda: create_borrow_and_return_tasks(['声光报警器'], borrow_list=['问拖车号'])).grid(row=1, column=2, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='负压循环洗厕所', width=0, command=lambda: create_borrow_and_return_tasks(['负压循环洗厕所'], borrow_list=['借洗厕所消毒水'], return_list=['还洗厕所消毒水'])).grid(row=1, column=3, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='吹散热器', width=0, command=lambda: create_borrow_and_return_tasks(['吹散热器'])).grid(row=2, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='叶片润滑', width=0, command=lambda: create_borrow_and_return_tasks(['叶片润滑'])).grid(row=2, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='换IDG', width=0, command=lambda: create_borrow_and_return_tasks(['IDG托架'])).grid(row=2, column=2, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='拿液压油', width=0, command=lambda: create_simple_task(['液压油柜'])).grid(row=2, column=3, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='倒废油', width=0, command=lambda: create_simple_task(['倒废油'])).grid(row=3, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='放水车', width=0, command=lambda: create_borrow_and_return_tasks(['放水车'])).grid(row=3, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(todo_frm, text='拿滑油', width=0, command=lambda: create_simple_task(['滑油柜'])).grid(row=3, column=2, padx=5, pady=2, sticky='nsew')

## guide中的内容按钮，使用grid布局，以免超出边界，目前是最多四列的排布
ttk.Button(guide_frm, text='本工具箱使用指南', width=0, command=lambda: webbrowser.open(UseGuide_url)).grid(row=0, column=0, padx=5, pady=2, sticky='nsew')
ttk.Button(guide_frm, text='各种手续操作指南', width=0, command=lambda: webbrowser.open(CaoZuoZhiNan_url)).grid(row=0, column=1, padx=5, pady=2, sticky='nsew')
ttk.Button(guide_frm, text='查件&件号', width=0, command=lambda: webbrowser.open(ChaJian_JianHao_url)).grid(row=0, column=2, padx=5, pady=2, sticky='nsew')
ttk.Button(guide_frm, text='工卡员指南', width=0, command=lambda: webbrowser.open(GongKaYuanZhiNan_url)).grid(row=0, column=3, padx=5, pady=2, sticky='nsew')
ttk.Button(guide_frm, text='工时自动化教学', width=0, command=lambda: os.startfile(auto_man_hours_document_path)).grid(row=1, column=0, padx=5, pady=2, sticky='nsew')


if __name__ == '__main__':
    # 开启时加载任务文件
    tasks_dict = load_json(todo_data_path)
    restore_tasks(tasks_dict)

    ## root的details_frm中创建A检信息展示区域
    show_check_info()

    ## 绑定窗口关闭事件，让窗口关闭时，保存数据
    ## 这里相当于把关闭GUI事件绑定了 保存数据库和关闭GUI，这里必须要destroy否则无法关闭GUI
    root.protocol('WM_DELETE_WINDOW', lambda: close_save(tasks_dict))

    root.mainloop()
