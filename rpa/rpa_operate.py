import csv
import io
import logging
import os
import queue
import re
import threading
import time
import uuid
from datetime import datetime, time as dt_time, timedelta
from typing import List, Dict, Optional

import ddddocr
import pandas as pd
import psutil
import pyautogui
import win32con
import win32file
from flask import Flask, request, jsonify
from pywinauto import Application, Desktop, timings, mouse, findwindows

from pypinyin import lazy_pinyin, Style
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys

from pojo.order import TradeRequest, StockRecord, Order, PendingOrder

log_dir = "../logs"  # 你可以替换成你想要的文件夹路径
if not os.path.exists(log_dir):
    os.makedirs(log_dir)
log_file = os.path.join(log_dir, "app.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),  # 写入文件
        logging.StreamHandler()  # 同时输出到控制台
    ]
)
log = logging.getLogger(__name__)

from pywinauto.controls.uiawrapper import UIAWrapper

def get_next_sibling(ctrl: UIAWrapper):
    """
    返回同级控件列表里，紧跟在 ctrl 后面的那个控件，没找到就返回 None。
    """
    parent = ctrl.parent()
    siblings = parent.children()
    try:
        idx = siblings.index(ctrl)
        return siblings[idx + 1]
    except (ValueError, IndexError):
        return None


def find_index(department_name, long_string):
    # 1. 先提取出营业部名称部分
    # 假设营业部名称部分在 "营业部名称" 之后，直到其他不相关内容
    # 你可以根据具体情况来调整分隔符
    start_index = long_string.find("营业部名称")  # 查找“营业部名称”部分的起始位置
    if start_index == -1:
        return -1  # 如果找不到营业部名称部分，返回-1

    # 截取出营业部名称部分
    department_part = long_string[start_index:]

    # 2. 找到实际的营业部名称部分，假设它们都在“营业部名称”之后
    # 通过分隔符来分割营业部名称，假设用空格作为分隔符
    department_list = department_part.split(" ")

    # 3. 将目标名称转为小写进行查找
    department_name = department_name.lower()

    # 4. 查找目标名称在列表中的位置
    for index, name in enumerate(department_list):
        if name.lower() == department_name:
            return index  # 返回第几个，注意是索引+1
    return -1  # 如果没有找到，返回-1


def first_order_switching(y: int) -> int:
    """
    元件到第一条（撤单使用）
    :param y:
    :return:
    """
    return y + 46

def order_switching(y: int) -> int:
    """
    下一条（撤单使用）
    :param y:
    :return:
    """
    return y + 17

def first_char_initial_upper(s: str) -> str | None:
    """
    取字符串 s 中第一个中文或英文字符的拼音首字母/字母，并大写返回，
    找不到返回 None。
    """
    for ch in s:
        # 英文字母：直接大写
        if re.match(r'[A-Za-z]', ch):
            return ch.upper()
        # 中文：取拼音首字母
        if '\u4e00' <= ch <= '\u9fff':
            return lazy_pinyin(ch, style=Style.FIRST_LETTER)[0].upper()
    return None

def cal(x: int, y: int) -> tuple[int, int]:
    return x + 499, y + 258


def smooth_move_and_scroll(target_x, target_y,
                           duration=1.0,
                           steps=50,
                           total_scroll=120):
    """
    在 duration 秒内分 steps 步，从当前位置平滑移动到 (target_x, target_y)，
    并在每步滚动 total_scroll/steps 的滚轮值。

    :param target_x: 目标 x 坐标（整数）
    :param target_y: 目标 y 坐标（整数）
    :param duration: 总移动时长（秒）
    :param steps:    拆分成多少步移动
    :param total_scroll: 整个过程累计滚动值（正数向上，负数向下）
    """
    start_x, start_y = pyautogui.position()
    dx = (target_x - start_x) / steps
    dy = (target_y - start_y) / steps
    scroll_step = total_scroll / steps
    step_duration = duration / steps

    for i in range(1, steps + 1):
        nx = start_x + dx * i
        ny = start_y + dy * i
        # 平滑小步移动
        pyautogui.moveTo(nx, ny, duration=step_duration)
        # 滚轮
        pyautogui.scroll(int(scroll_step))
        # （可选）微小延迟加强平滑感
        # time.sleep(0.001)

def find_order_record_cn(data_list: List[Dict], target: Order) -> Optional[Dict]:
    for rec in data_list:
        if (rec.get("证券代码") == target.symbol
            and rec.get("方向")     == target.operation
            and rec.get("合同编号") == target.contract_number):
            return rec
    return None


def flush_cache(path: str):
    """
    刷新 Windows 文件系统缓存，确保下次读取获取到最新的文件内容。
    使用 CreateFile 打开带读写和共享模式的句柄，并调用 FlushFileBuffers。
    """
    try:
        handle = win32file.CreateFile(
            path,
            win32con.GENERIC_READ | win32con.GENERIC_WRITE,
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
            None,
            win32file.OPEN_EXISTING,
            0,
            None
        )
        win32file.FlushFileBuffers(handle)
        win32file.CloseHandle(handle)
        log.info("🔄 Windows 文件系统缓存已刷新")
    except Exception as e:
        log.warning(f"⚠️ 刷新文件系统缓存失败：{e}")


class RPAOperate:
    def __init__(self, config):

        self.app = None
        self.config = {
            'sleep_waiting_trade': 120,
            'exe_path': r"D:\360Downloads\Software\同花顺\xiadan.exe",
            'trade_account': config.get("trade_account", ""), # 交易账号
            'trade_password': config.get("trade_password", ""), # 交易密码
            'brokerage': config.get("brokerage", ""), # 券商
            'opening_area': config.get("opening_area", ""), # 开户地区（可为空）,0:普通，1:融资融券
            'sales_department': config.get("sales_department", ""), # 营业厅（可为空）
            'business': config.get("business", ""), # 0:综合业务，1:信用业务
        }

        self.main_wnd = None
        self.login_window = None
        self.second_login_wnd = None

        # 订单队列
        # 整体rpa操作执行队列
        # self.queue: queue.PriorityQueue = queue.PriorityQueue()
        self.high_priority_queue = queue.Queue()
        self.low_priority_queue = queue.Queue()
        self.worker_thread = None  # 订单处理线程，提前声明一个属性以便管理线程

        self.grid_height = None
        self.MAX_RECORD_NUM = None

        # 用于文件更新操作
        self._last_mtime = None
        # 账户余额
        self.funds = None

    def connect(self):
        # 0: 已登录 1: 登录成功 10: 其他原因登录失败
        log_status = 10
        login_times = 0
        while login_times < 3:
            login_times += 1
            log.info(
                f"第{login_times}次连接客户端: <{self.config['exe_path']}>[{self.config['trade_account']}:{self.config['brokerage']}]")
            log_status = self.__login()
            if log_status == 1 or log_status == 0:
                log.info(
                    f"第{login_times}次连接成功: <{self.config['exe_path']}>[{self.config['trade_account']}:{self.config['brokerage']}]")
                break
            log.error(
                f"第{login_times}次连接失败: <{self.config['exe_path']}>[{self.config['trade_account']}:{self.config['brokerage']}]")
            if log_status == 2 or log_status == 3:
                break
            time.sleep(60)
        if log_status != 0 and log_status != 1:
            message = ""
            if log_status == 2:
                message = "密码有误"
            elif log_status == 3:
                message = "客户号无效"
            elif log_status == 10:
                message = "关联失败，请联系客服"

            log.warning(
                f"RPA 登录失败 账户: <{self.config['exe_path']}>[{self.config['trade_account']}:{self.config['brokerage']}] 原因：{message}")
            return {'success': False, 'message': message, 'log_status': log_status}

        log.info(f"正在连接客户端: {self.config['exe_path']}")
        self.app = Application().connect(
            path=self.config["exe_path"], timeout=10)
        log.info(f"已连接到: {self.config['exe_path']}")
        self.main_wnd = self.app.window(
            title="网上股票交易系统5.0")  # self.main_wnd
        self.main_wnd.set_focus()

        # 窗口最大化
        if not self.main_wnd.was_maximized():
            self.main_wnd.maximize()

        # 获取委托单列表高度
        self.__select_menu(['市价委托', '卖出'])
        send_keys('{VK_F8}')  # 点击委托选项卡
        grid_zone = self.main_wnd.window(control_id=0x417, class_name='CVirtualGridCtrl')
        self.grid_height = grid_zone.rectangle().height() - 40
        # 一条记录的高度
        ONE_RECORD_HEIGHT = 16
        # 委托单一共可以显示记录的条数
        self.MAX_RECORD_NUM = self.grid_height // ONE_RECORD_HEIGHT
        log.info(f"委托单列表最多显示记录条数: {self.MAX_RECORD_NUM}")

        return {'success': True, 'message': 'RPA 登陆成功', 'log_status': log_status}

    def __login(self):
        try:
            main = Application(backend="uia").connect(best_match="网上股票交易系统")
            self.main_wnd = main.window(title_re=".*网上股票交易系统.*")
            tool_bar = self.main_wnd.child_window(control_type="ToolBar")
            try:
                combobox = tool_bar.child_window(control_type="ComboBox", found_index=2)
                # 获取控件所在的进程 ID
                process_id = combobox.process_id()
                print(f"Process ID of ComboBox: {process_id}")
                combobox.click_input()
                time.sleep(1)
                # 连接到正在运行的应用程序
                app = Application(backend="uia").connect(process=process_id)
                # 2. 找到所有顶层窗口的句柄
                hwnds = findwindows.find_windows(process=process_id)
                if not hwnds:
                    raise RuntimeError("⚠️ 没有找到任何窗口")
                # 3. 取第一个句柄，构造 Specification
                main_spec = app.window(handle=hwnds[0])
                pattern = f".*{re.escape(self.config['brokerage'])}.*"
                # 这一步要么拿到控件，要么抛 ElementNotFoundError
                ctrl = main_spec.child_window(
                    control_type="ListItem",
                    title_re=pattern
                ).wrapper_object()
                ctrl.click_input()
                return 0
            except:
                return self.secondary_login()
        except Exception as e:
            res = self.before_login()
            time.sleep(10)
            if res:
                return self.secondary_login()
            else:
                return 10

    def before_login(self):
        """
        前置登录:登录模拟账号
        :return: bool
        """
        try:
            app = Application(backend="uia").start(
                self.config['exe_path'], timeout=10
            )
            time.sleep(15)
            TARGET_EXE = "xiadan.exe"
            BACKEND = "uia"

            timings.window_find_timeout = 15.0
            timings.after_click_wait = 1.0

            # 1. attach 进程
            pid = next(p.info['pid'] for p in psutil.process_iter(['pid', 'name'])
                       if p.info['name'] == TARGET_EXE)
            app = Application(backend=BACKEND).connect(process=pid)
            desktop = Desktop(backend=BACKEND)

            # 2. 顶层窗口
            wins = desktop.windows(process=pid, top_level_only=True, visible_only=False)
            dlg2 = app.window(handle=wins[1].handle)
            self.login_window = dlg2
            rect = dlg2.rectangle()
            x, y = cal(rect.left, rect.top)
            mouse.click(button='left', coords=(x, y))

            # TODO
            try:
                dlg2.child_window(control_type="Button", best_match="选择券商").click_input()
            except:
                # self.login_window.child_window(control_type="Button", title="")
                select = dlg2.child_window(control_type="ComboBox", found_index=0)
                select_rect = select.rectangle()
                select_mid = select_rect.mid_point()
                mouse.click(button='left', coords=(int(select_mid.x - 30), int(select_mid.y)))

            finally:
                doc_spec = dlg2.child_window(control_type="Document", found_index=0)

                # 3. 点击首字母索引并根据该元素位置向下滚动
                letter_ctrl = doc_spec.child_window(
                    control_type="Text",
                    best_match=first_char_initial_upper("模拟炒股")
                )
                letter_ctrl.click_input()
                time.sleep(0.3)

                # ——— 根据 letter_ctrl 位置滑动 ———
                # 取出它的矩形中心点
                rect = letter_ctrl.rectangle()
                mid = rect.mid_point()
                print(mid.x, mid.y)

                time.sleep(0.5)

                # 4. → 可能找不到目标券商 Text，这里加 try‑except
                try:
                    target_brokerage = doc_spec.child_window(
                        control_type="Text",
                        best_match="模拟炒股"
                    )
                    tb_wrapper = target_brokerage.wrapper_object()
                except Exception as e:
                    print(f"❌ 未找到券商 «{self.config['brokerage']}»：{e}")
                    return False  # 结束，返回失败

                # 5. 让鼠标飘过去（可选）
                center = tb_wrapper.rectangle().mid_point()
                mouse.move(coords=(center.x, center.y))
                time.sleep(0.3)

                # 6. 同一级兄弟里找同一行按钮；如果没找到再捕获一次
                broker_y = center.y
                try:
                    siblings = tb_wrapper.parent().children()
                    bind_btn = next(
                        sib for sib in siblings
                        if sib.element_info.control_type == "Button"
                        and sib.window_text() == "绑定已有账户"
                        and abs(sib.rectangle().mid_point().y - broker_y) < 3
                    )
                except StopIteration:
                    print(f"❌ «{self.config['brokerage']}» 行内未找到『绑定已有账户』按钮")
                    return False  # 结束，返回失败

                bind_btn.move_mouse_input()
                bind_btn.click_input()
                time.sleep(3)
                try:
                    # —— 2. 定位到两个 Image 元件 ——
                    verification_code = dlg2.child_window(title="交易密码:", found_index=1, control_type="Image")
                    # 3. 截图并保存到内存
                    pil_img = verification_code.capture_as_image()  # 返回 PIL.Image
                    buf = io.BytesIO()
                    pil_img.save(buf, format="PNG")
                    img_bytes = buf.getvalue()

                    # 4. 用 ddddocr 识别
                    ocr = ddddocr.DdddOcr()
                    res = ocr.classification(img_bytes)

                    count = len(dlg2.children(control_type="Edit"))
                    # found_index 从 0 开始，到 count-1
                    dlg2.child_window(control_type="Edit", found_index=count).click_input()
                    log.info(f"识别结果：{res}")
                    # 持续 1.5 秒，不停地发 BACKSPACE
                    end_time = time.time() + 0.5
                    while time.time() < end_time:
                        send_keys('{BACKSPACE}', pause=0)  # pause=0 加速
                    time.sleep(0.3)
                    send_keys(res)
                finally:
                    time.sleep(1)
                    dlg2.child_window(control_type="Button", title="登录").click_input()
                    return True
        except Exception as e:
            print(f"❌ 登录前置步骤发生异常：{e}")
            return False  # 出现异常，返回失败

    def secondary_login(self):
        try:
            main = Application(backend="uia").connect(best_match="网上股票交易系统")
            self.main_wnd = main.window(title_re=".*网上股票交易系统.*")
            self.main_wnd.child_window(control_type="Button", best_match="添加").click_input()
            res = self.real_login()
            time.sleep(5)
            if res:
                return self.login_edit()
            else:
                return 10
        except Exception as e:
            print(f"❌ secondary_login 失败: {e}")
            return 10  # 失败，返回False

    def real_login(self):
        """
        登录用户账号
        :return: bool
        """
        try:
            second_login_window = self.main_wnd.child_window(control_type="Pane", found_index=1)
            self.second_login_wnd = second_login_window
            second_login_window.child_window(control_type="Button", best_match="选择券商").click_input()
            doc_spec = second_login_window.child_window(control_type="Document", found_index=0)

            # 3. 点击首字母索引并根据该元素位置向下滚动
            letter_ctrl = doc_spec.child_window(
                control_type="Text",
                best_match=first_char_initial_upper(self.config['brokerage'])
            )
            letter_ctrl.click_input()
            time.sleep(0.3)

            # ——— 根据 letter_ctrl 位置滑动 ———
            # 取出它的矩形中心点
            rect = letter_ctrl.rectangle()
            mid = rect.mid_point()
            print(mid.x, mid.y)

            time.sleep(0.5)

            # 4. → 可能找不到目标券商 Text，这里加 try‑except
            try:
                target_brokerage = doc_spec.child_window(
                    control_type="Text",
                    best_match=self.config['brokerage']
                )
                tb_wrapper = target_brokerage.wrapper_object()
            except Exception as e:
                x = int(mid.x)
                y = int(mid.y + 50)
                smooth_move_and_scroll(x, y, 0.1, 1, -500)

                try:
                    target_brokerage = doc_spec.child_window(
                        control_type="Text",
                        best_match=self.config['brokerage']
                    )
                    tb_wrapper = target_brokerage.wrapper_object()
                except Exception as e:
                    print(f"❌ 未找到券商 «{self.config['brokerage']}»：{e}")
                    return False  # 未找到券商，返回False

            # 5. 让鼠标飘过去（可选）
            center = tb_wrapper.rectangle().mid_point()
            mouse.move(coords=(center.x, center.y))
            time.sleep(0.3)

            # 6. 同一级兄弟里找同一行按钮；如果没找到再捕获一次
            broker_y = center.y
            try:
                siblings = tb_wrapper.parent().children()
                bind_btn = next(
                    sib for sib in siblings
                    if sib.element_info.control_type == "Button"
                    and sib.window_text() == "绑定已有账户"
                    and abs(sib.rectangle().mid_point().y - broker_y) < 3
                )
            except StopIteration:
                print(f"❌ «{self.config['brokerage']}» 行内未找到『绑定已有账户』按钮")
                return False  # 找不到按钮，返回False

            bind_btn.move_mouse_input()
            bind_btn.click_input()
            time.sleep(2)
            try:
                doc = self.second_login_wnd.child_window(control_type="Document", found_index=0)

                if self.config['business']:
                    if int(self.config['business']) == 0:
                        doc.child_window(control_type="Text", title_re=".*综合业务.*").click_input()
                    elif int(self.config['business']) == 1:
                        doc.child_window(control_type="Text", title_re=".*信用业务.*").click_input()
                    else:
                        raise Exception("错误选择")
                elif self.config['opening_area'] and self.config['sales_department']:
                    if int(self.config['opening_area']) == 0:
                        doc.child_window(control_type="Text", title_re=".*普通.*").click_input()
                        title = doc.window_text()
                        print(f"控件标题: {title}")
                        index = find_index(self.config['sales_department'], title)
                        current_pos = pyautogui.position()
                        new_x = current_pos[0] + 264
                        new_y = current_pos[1]
                        mouse.click(coords=(new_x, new_y))
                        log.info(index)
                        for _ in range(index // 3):
                            pyautogui.scroll(-118)
                        doc.child_window(control_type="Text", best_match=self.config['sales_department']).click_input()
                        doc.child_window(control_type="Button", best_match="添加").click_input()

                    elif int(self.config['opening_area']) == 1:
                        all_items = doc.descendants(control_type="Text", title="融资融券")
                        if all_items:
                            last_item = all_items[-1]
                            print("最后一个同名控件句柄：", last_item.handle)
                            last_item.click_input()  # 点击融资融券
                        else:
                            print("没找到任何“融资融券”")

                        title = doc.window_text()
                        print(f"控件标题: {title}")
                        index = find_index(self.config['sales_department'], title)
                        current_pos = pyautogui.position()
                        new_x = current_pos[0] + 264
                        new_y = current_pos[1]
                        mouse.click(coords=(new_x, new_y))
                        log.info(index)
                        for _ in range(index // 3):
                            pyautogui.scroll(-118)
                        doc.child_window(control_type="Text", best_match=self.config['sales_department']).click_input()
                        doc.child_window(control_type="Button", best_match="添加").click_input()

            except Exception as e:
                log.error(e)
                return False  # 异常处理，返回False

            return True  # 成功，返回True
        except:
            return False

    def login_edit(self):
        try:
            account = self.second_login_wnd.child_window(control_type="Edit", found_index=0)
            account.click_input()
            send_keys('{BACKSPACE}')
            send_keys(self.config['trade_account'])
            time.sleep(0.3)
            send_keys("{TAB}")
            time.sleep(0.3)
            send_keys(self.config['trade_password'])
            send_keys("{TAB}")
            time.sleep(0.3)
            try:
                # —— 2. 定位到两Image 元件 ——
                verification_code = self.second_login_wnd.child_window(title="交易密码:", found_index=1,
                                                                       control_type="Image")
                pil_img = verification_code.capture_as_image()  # 返回 PIL.Image
                buf = io.BytesIO()
                pil_img.save(buf, format="PNG")
                img_bytes = buf.getvalue()

                # 用 ddddocr 识别
                ocr = ddddocr.DdddOcr()
                res = ocr.classification(img_bytes)

                log.info(f"识别结果：{res}")
                end_time = time.time() + 0.5
                while time.time() < end_time:
                    send_keys('{BACKSPACE}', pause=0)  # pause=0 加速
                time.sleep(0.3)
                send_keys(res)
                # 登录按钮点击
                self.second_login_wnd.child_window(control_type="Button", title="登录").click_input()
                time.sleep(1)
                try:
                    result = self.login_result()
                    if "密码有误" in result:
                        return 2
                    elif "客户号" in result:
                        return 3
                    else:
                        return 10
                except:
                    self.main_wnd.maximize()
                    return 1
            except:
                # 登录按钮点击
                self.second_login_wnd.child_window(control_type="Button", title="登录").click_input()
                time.sleep(1)
                try:
                    result = self.login_result()
                    if "密码有误" in result:
                        return 2
                    elif "客户号" in result:
                        return 3
                    else:
                        return 10
                except:
                    self.main_wnd.maximize()
                    return 1

        except Exception as e:
            log.error(e)
            return 10  # 失败，返回False

    def login_result(self) -> str:
        parent = self.main_wnd.child_window(control_type="Pane", found_index=1)
        res = parent.child_window(control_type="Image", auto_id="1004")
        result = res.element_info.name
        print(result)
        self.main_wnd.child_window(control_type="Button", title="确定", found_index=0).click_input()
        self.second_login_wnd.close()
        return result

    def __select_menu(self, path):
        """ 点击左边菜单 """
        if r"网上股票" not in self.main_wnd.window_text():
            self.main_wnd.set_focus()
            send_keys("{ENTER}")
        self.__get_left_menus_handle().get_item(path).click()

    def __get_left_menus_handle(self):
        while True:
            try:
                handle = ""
                try:
                    handle = self.main_wnd.window(
                        control_id=0x81, class_name='SysTreeView32')
                except:
                    time.sleep(self.config["sleepC"])
                    handle = self.main_wnd.window(
                        control_id=0x81, class_name='SysTreeView32')
                handle.wait('ready', 2)
                return handle
            except Exception as ex:
                log.info(ex)
                pass


    def add_task(self, trade_request, high_priority: bool = False):
        """ 向任务队列添加任务 """
        try:
            q = self.high_priority_queue if high_priority else self.low_priority_queue
            log.info(f"即将入队任务: {trade_request}")
            q.put(trade_request)  # 阻塞队列，自动等待直到队列有空位
            log.info(f"添加 {'高优先' if high_priority else '低优先'} 任务: {trade_request};")
            return True
        except Exception as e:
            log.error(f"添加任务失败: {e}")
            return False

    def add_tasks(self, trade_requests, high_priority: bool = False):
        """
        一次性批量添加多个任务到队列
        """
        for trade_request in trade_requests:
            self.add_task(trade_request, high_priority)

    def handle_task(self):

        # allowed_periods = [
        #     (dt_time(9, 30), dt_time(11, 30)),  # 上午段：09:30 ~ 11:30
        #     (dt_time(13, 0), dt_time(15, 0))  # 下午段：13:00 ~ 15:00
        # ]
        allowed_periods = [
            (dt_time.min, dt_time.max)
        ]
        # 下次调用 get_funds 的时间，初始设置为当前时间，即一开始就尝试调用
        next_funds_call = datetime.now()
        while True:
            current_time = datetime.now().time()
            in_allowed_period = any(start <= current_time <= end for start, end in allowed_periods)

            if not in_allowed_period:
                log.info("当前时间不在任务受理时间段内，小憩一下，梦里也别忘了资金呢...")
                time.sleep(30)  # 不在允许时间段内则休眠 30 秒
                continue

            # 如果到了资金查询时间，就顺便查下资金情况，给钱钱来个“打卡”
            if datetime.now() >= next_funds_call:
                log.info("执行资金查询（get_funds）接口")
                self.get_funds()
                next_funds_call = datetime.now() + timedelta(minutes=15)

            task = self.get_next_task()  # 自动阻塞等待任务
            if not task:
                continue  # 如果没有任务则跳过本次循环

            max_retries = 3
            for attempt in range(1, max_retries + 1):
                try:
                    self.get_funds()
                    log.info(f"执行交易任务 {task}，第 {attempt} 次尝试")

                    result = self.operation(task)

                    # 检查订单结果是否不为空
                    if result:
                        log.info(f"交易任务 {task} 成功完成")
                        break  # 成功处理，跳出重试循环
                    elif not result:
                        log.info(f"交易任务 {task} 失败")
                        break
                    else:
                        raise Exception("order_result 返回空结果，继续重试")
                except Exception as e:
                    log.error(f"交易任务 {task} 第 {attempt} 次失败: {e}")
                    if attempt < max_retries:
                        time.sleep(1)  # 重试前的延迟
                    else:
                        log.error(f"任务 {task} 连续失败，放弃处理")

    def start_worker(self):
        t = threading.Thread(target=self.handle_task, daemon=True)
        t.start()
        self.worker_thread = t
        log.info("交易任务线程已启动")

    def get_next_task(self):
        """ 从队列中取任务，如果没有任务就阻塞等待 """
        try:
            # 优先取高优先级任务
            task = self.high_priority_queue.get(timeout=1)  # 阻塞等待直到有任务
            return task
        except queue.Empty:
            # 如果高优先级队列为空，则从低优先级队列获取任务
            try:
                task = self.low_priority_queue.get(timeout=1)  # 阻塞等待直到有任务
                return task
            except queue.Empty:
                return None  # 两个队列都没有任务时，返回 None

    def operation(self, task: Optional[TradeRequest] = None, order: Optional[Order] = None) -> bool:
        """
        执行交易操作：
          0 -> 买入（只需传入 task）
          1 -> 卖出（只需传入 task）
          2 -> 撤单（只需传入 order）
        参数缺失或未知操作码均返回 False。
        """
        # 聚焦主窗口
        self.main_wnd.set_focus()

        # 撤单，只看 order
        if order is not None:
            send_keys('{F3}')
            return self.cancel_task(order)

        # 买入/卖出，需要 task
        if task is None:
            log.error("❌ 缺少 TradeRequest，无法执行买入/卖出")
            return False

        match task.operation:
            case 0:
                send_keys('{F1}')
                return self.trade(task, "买入")
            case 1:
                send_keys('{F2}')
                return self.trade(task, "卖出")
            case _:
                log.error(f"❌ 未知操作码：{task.operation}")
                return False

    def cancel_task(self, order: Order):
        self.save()
        data_list = self.deal_with_xsl()
        time.sleep(1)

        # 用 enumerate 把索引也带上，start=1 表示第一条就是第 1 条
        index, matched = next(
            (
                (i, rec)
                for i, rec in enumerate(data_list, start=1)
                if rec.get('证券代码') == order.symbol
                   and rec.get('操作') == order.operation
                   and rec.get('合同编号') == order.contract_number
            ),
            (None, None)
        )

        if matched is None:
            log.error("❌ 找不到对应记录")
            return False
        else:
            print(f"✅ 匹配到第 {index} 条数据：{matched}")
            # 如果后面要根据 index 操作 UI，比如第几行点击，就直接用 index

        transition = self.main_wnd.child_window(control_type="Button", title="撤最后(G)")
        transition_rect = transition.rectangle()
        transition_mid = transition_rect.mid_point()
        y = first_order_switching(transition_mid.y)
        for x in range(index - 1):
            y = order_switching(y)
        mouse.double_click(button='left', coords=(transition_mid.x, y))
        self.main_wnd.child_window(control_type="Button", title="是(Y)").click_input()
        return True


    def history_orders(self):
        self.main_wnd.maximize()
        self.main_wnd.set_focus()
        send_keys('{F4}')
        self.save()
        data_list = self.deal_with_xsl()
        return [StockRecord.from_dict(d) for d in data_list]


    def trade(self, task: TradeRequest, opera: str):
        main = Application(backend="uia").connect(best_match="网上股票交易系统")
        self.main_wnd = main.window(title_re=".*网上股票交易系统.*")
        if self.funds < (task.price * task.volume) and task.operation == 0:
            log.warning("可用余额不足")
            return False
        try:
            code_edit = self.main_wnd.child_window(
                control_type="Edit", found_index=0)  # 股票代码输入框
            price_edit = self.main_wnd.child_window(
                control_type="Edit", found_index=1)  # 价格输入框
            quantity_edit = self.main_wnd.child_window(
                control_type="Edit", found_index=2)  # 数量输入框
        except Exception as e:
            log.error(f"定位输入框失败: {str(e)}")
            return False

        try:
            # 清空并输入股票代码（直接使用set_text方法更可靠）
            code_edit.set_focus()
            # for i in range(10):
            #     send_keys('{BACKSPACE}')
            code_edit.click_input(double=True)
            send_keys('{BACKSPACE}')
            time.sleep(0.2)
            send_keys(task.symbol)
            time.sleep(0.2)

            send_keys('{TAB}')
            time.sleep(0.2)
            price_edit.click_input(double=True)
            time.sleep(0.2)
            send_keys(str(task.price))

            time.sleep(0.2)

            # 输入数量
            send_keys('{TAB}')
            time.sleep(1)
            quantity_edit.click_input(double=True)
            time.sleep(0.2)
            send_keys(str(task.volume))

            # 查找并点击买入按钮（根据实际按钮名称修改）
            buy_btn = self.main_wnd.child_window(
                best_match=opera, control_type="Button")
            buy_btn.click_input()

            try:
                time.sleep(2)
                send_keys("{Y}")
                send_keys("{Y}")
                # self.main_wnd.print_control_identifiers()
                failed_img_spec = self.main_wnd.child_window(control_type="Image", title_re=".*失败.*")
                if failed_img_spec.wait("exists", timeout=1):
                    self.main_wnd.child_window(title="确定", control_type="Button").click_input()
                    return False
            except Exception as e:
                # TODO 交易成功待测试
                time.sleep(1)
                success_img_spec = self.main_wnd.child_window(control_type="Image", title_re=".*成功.*")
                if success_img_spec.exists(timeout=1):
                    self.main_wnd.child_window(title="确定", control_type="Button").click_input()
                    return True
                else:
                    return False
        except Exception as e:
            log.error(f"交易操作执行失败: {str(e)}")
            return False


    def save(self):
        self.main_wnd.maximize()
        self.main_wnd.set_focus()

        send_keys('^s')
        try:
            time.sleep(3)
            verification_code = self.main_wnd.child_window(control_type="Image", title_re=".*正在保存数据.*", found_index=1)
            # 3. 截图并保存到内存
            pil_img = verification_code.capture_as_image()  # 返回 PIL.Image
            buf = io.BytesIO()
            pil_img.save(buf, format="PNG")
            img_bytes = buf.getvalue()

            # 4. 用 ddddocr 识别
            ocr = ddddocr.DdddOcr()
            res = ocr.classification(img_bytes)

            self.main_wnd.child_window(control_type="Edit", title="提示").click_input()
            send_keys(str(res))

            self.main_wnd.child_window(control_type="Button", title="确定").click_input()
        except Exception:
            log.warning("没有找到验证码，忽略")
        finally:
            export = self.main_wnd.child_window(control_type="Window", title_re="另存为")
            # 获取当前用户主目录
            home_dir = os.path.expanduser('~')
            # 根据系统决定使用 'Desktop' 或 '桌面' 文件夹名称
            desktop_folder = 'Desktop'  # 如果是中文系统则可能为 '桌面'
            desktop_path = os.path.join(home_dir, desktop_folder)

            # 拼接出桌面上的 futures 文件夹路径
            stock_folder = os.path.join(desktop_path, 'stock')

            # 检查并创建
            if not os.path.exists(stock_folder):
                os.makedirs(stock_folder)
            # TODO 不同的操作系统可能不太一样
            progress = export.child_window(best_match="Progress")
            # 获取控件的矩形区域
            first_rect = progress.rectangle()
            # second_rect = export.rectangle()
            print(f"Progress 的 rlbt 为：{first_rect}")  # rect 格式大致为 Rect(left, top, right, bottom)
            # 计算两个控件中间位置的坐标，并取整
            middle_x = int((first_rect.right + first_rect.left) / 2)
            middle_y = int((first_rect.bottom + first_rect.top) / 2)
            # time.sleep(3)
            # 点击中间位置
            mouse.click(button='left', coords=(middle_x, middle_y))

            send_keys(stock_folder)
            time.sleep(1)
            send_keys('{ENTER}')

            export.child_window(control_type="Button", title_re=".*保存.*").click_input()
            try:
                time.sleep(1)
                export.child_window(control_type="Button", title_re=".*是.*").click_input()
            except Exception as e:
                log.warning("第一次保存")

    def deal_with_xsl(self):

        # 1）定位桌面 stock 文件夹
        home_dir = os.path.expanduser('~')
        desktop_name = 'Desktop'
        stock_folder = os.path.join(home_dir, desktop_name, 'stock')
        os.makedirs(stock_folder, exist_ok=True)
        orig_path = os.path.join(stock_folder, 'table.xls')
        log.info(f"读取文件路径：{orig_path}")

        # 2）检查文件是否存在
        if not os.path.exists(orig_path):
            log.error(f"❌ 文件不存在：{orig_path}")
            return None

        # 3）获取当前文件状态并强制刷新
        current_mtime = os.path.getmtime(orig_path)
        current_size = os.path.getsize(orig_path)
        flush_cache(orig_path)

        # 4）确保文件处于稳定状态（不在写入过程中）
        stable_retries = 5
        for i in range(stable_retries):
            time.sleep(0.2)  # 短暂等待确保文件写完
            flush_cache(orig_path)
            new_mtime = os.path.getmtime(orig_path)
            new_size = os.path.getsize(orig_path)

            # 如果两次检查大小和修改时间都相同，文件应该稳定了
            if new_mtime == current_mtime and new_size == current_size:
                log.info(f"✅ 文件稳定，mtime={new_mtime}, size={new_size}")
                break
            log.info(f"⏳ 第{i + 1}次检测文件正在变化，等待稳定")
            current_mtime = new_mtime
            current_size = new_size

        # 5）使用临时副本读取，添加随机后缀避免命名冲突
        temp_name = f"table_{uuid.uuid4().hex}.xls"
        temp_path = os.path.join(stock_folder, temp_name)
        try:
            # 使用低级复制函数，确保不使用系统缓存
            with open(orig_path, 'rb') as fsrc:
                with open(temp_path, 'wb') as fdst:
                    fdst.write(fsrc.read())

            # 确保复制完成
            os.sync() if hasattr(os, 'sync') else None  # 在支持的系统上同步文件系统
            read_path = temp_path
            log.info(f"✅ 创建临时副本成功：{temp_path}")
        except Exception as e:
            log.warning(f"⚠️ 复制到临时文件失败，直接读取原始文件：{e}")
            read_path = orig_path

        # 6）开始尝试多种方式读取文件
        df = None

        # 尝试直接读取文件头来确定编码和格式
        try:
            with open(read_path, 'rb') as f:
                header = f.read(100)
                log.info(f"文件头部字节: {header}")
        except Exception as e:
            log.warning(f"⚠️ 无法读取文件头: {e}")

        # 首先尝试 CSV+GBK+制表符 读取
        try:
            with open(read_path, newline='', encoding='gbk') as csvfile:
                reader = csv.reader(csvfile, delimiter='\t')
                rows = list(reader)
            if rows:
                df = pd.DataFrame(rows[1:], columns=rows[0])
                log.info("✅ 成功用 CSV+GBK+制表符 读取并解析文件")
        except UnicodeDecodeError:
            # 尝试 CSV+UTF-8+制表符
            try:
                with open(read_path, newline='', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile, delimiter='\t')
                    rows = list(reader)
                if rows:
                    df = pd.DataFrame(rows[1:], columns=rows[0])
                    log.info("✅ 成功用 CSV+UTF-8+制表符 读取并解析文件")
            except Exception as e:
                log.warning(f"⚠️ 用 CSV+制表符 读取失败(UTF-8)：{e}")

                # 尝试 CSV+GBK+逗号
                try:
                    with open(read_path, newline='', encoding='gbk') as csvfile:
                        reader = csv.reader(csvfile, delimiter=',')
                        rows = list(reader)
                    if rows:
                        df = pd.DataFrame(rows[1:], columns=rows[0])
                        log.info("✅ 成功用 CSV+GBK+逗号 读取并解析文件")
                except Exception as e:
                    log.warning(f"⚠️ 用 CSV+逗号 读取失败(GBK)：{e}")
        except Exception as e:
            log.warning(f"⚠️ 用 CSV 读取失败(GBK)：{e}")

        # 如果还是失败，尝试Excel格式读取
        if df is None:
            try:
                df = pd.read_excel(read_path, engine='openpyxl')
                log.info("✅ 成功读取 .xlsx 文件")
            except Exception as e:
                log.warning(f"⚠️ 读取 .xlsx 失败：{e}")
                try:
                    df = pd.read_excel(read_path, engine='xlrd')
                    log.info("✅ 成功读取 .xls 文件")
                except Exception as e:
                    log.error(f"❌ 读取 .xls 失败：{e}")

                    # 最后一次尝试：通用文本读取
                    try:
                        df = pd.read_csv(read_path, sep=None, engine='python', encoding='gbk')
                        log.info("✅ 成功用pandas自动检测分隔符模式读取")
                    except Exception as e:
                        log.error(f"❌ 所有读取方法都失败：{e}")
                        log.error("文件可能已损坏或格式不正确")

        # 7）清理临时文件
        if read_path != orig_path:
            try:
                os.remove(read_path)
                log.info(f"✅ 清理临时文件成功")
            except Exception as e:
                log.warning(f"⚠️ 删除临时文件失败：{e}")

        # 8）返回结果
        if df is not None:
            # 确保数据框不为空
            if df.empty:
                log.warning("⚠️ 数据框为空")
                return []

            # 转换为字典列表
            try:
                data_list = df.to_dict(orient='records')
                log.info(f"✅ 数据已封装为字典列表，共{len(data_list)}条")
                sample_size = min(3, len(data_list))
                if sample_size > 0:
                    log.info(f"示例前{sample_size}条：")
                    for rec in data_list[:sample_size]:
                        log.info(rec)
                return data_list
            except Exception as e:
                log.error(f"❌ 转换为字典列表失败：{e}")
                return None
        else:
            log.error("❌ 无法读取文件内容，返回 None")
            return None

    def pending_orders(self):
        self.main_wnd.set_focus()
        send_keys("{F3}")
        self.save()
        data_list = self.deal_with_xsl()
        result: List[PendingOrder] = []
        for item in data_list:
            order = PendingOrder(
                market=item.get('交易市场', ''),
                contract_number=item.get('合同编号', ''),
                remark=item.get('备注', ''),
                order_price=float(item.get('委托价格', 0) or 0),
                order_quantity=int(item.get('委托数量', 0) or 0),
                order_time=item.get('委托时间', ''),
                average_price=float(item.get('成交均价', 0) or 0),
                trade_quantity=int(item.get('成交数量', 0) or 0),
                cancel_quantity=int(item.get('撤消数量', 0) or 0),
                operation=item.get('操作', ''),
                symbol=item.get('证券代码', ''),
                security_name=item.get('证券名称', '')
            )
            result.append(order)
        return result

    def get_funds(self):
        self.main_wnd.set_focus()
        send_keys("{F4}")
        num1 = self.main_wnd.window(
            control_id=0x3F4, class_name='Static')
        time.sleep(0.1)
        num2 = self.main_wnd.window(
            control_id=0x3F5, class_name='Static')
        time.sleep(0.1)
        num3 = self.main_wnd.window(
            control_id=0x3F8, class_name='Static')
        time.sleep(0.1)
        num4 = self.main_wnd.window(
            control_id=0x3F9, class_name='Static')
        time.sleep(0.1)
        num5 = self.main_wnd.window(
            control_id=0x3F6, class_name='Static')
        time.sleep(0.1)
        num6 = self.main_wnd.window(
            control_id=0x3F7, class_name='Static')
        time.sleep(0.1)
        num7 = self.main_wnd.window(
            control_id=0x403, class_name='Static')
        time.sleep(0.1)
        df = [{
            "资金金额": float(num1.texts()[0]),
            "冻结金额": float(num2.texts()[0]),
            "可用金额": float(num3.texts()[0]),
            "可取金额": float(num4.texts()[0]),
            "股票市值": float(num5.texts()[0]),
            "总资产": float(num6.texts()[0]),
            "持仓盈亏": float(num7.texts()[0])
        }]
        self.funds = float(num3.texts()[0])
        return float(num3.texts()[0])




if __name__ == '__main__':
    pass