import csv
import subprocess
import threading
import traceback
import queue  # 替代 multiprocessing.Queue
from typing import Any

import re  # 导入正则表达式模块
import schedule
import time
from datetime import datetime, time as dt_time, timedelta

import pywinauto
import pytz
from pywinauto import timings, mouse
from pywinauto.keyboard import send_keys

import os
import logging

from pojo.order import TradeRequest, HistoryOrder

from flask import Flask, request, jsonify



# 存储日志配置
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


def get_bottom_rect(l, t, r, b, offset_top=21, offset_bottom=26):
    """
    根据给定的偏移量调整矩形的上边界和下边界

    参数：
        l: 左边界
        t: 上边界
        r: 右边界
        b: 下边界
        offset_top: 上边界的偏移量，默认值 21
        offset_bottom: 下边界的偏移量，默认值 26

    返回：
        调整后的矩形坐标 (l, t+offset_top, r, b+offset_bottom)
    """
    new_t = t + offset_top
    new_b = b + offset_bottom
    return l, new_t, r, new_b


class RpaOperator:

    def __init__(self, config):
        self.app = None
        self.config = {
            'sleep_start': 5,
            'sleep_operation': 0.5,
            'sleep_login': 8,
            'sleep_waiting_trade': 120,
            'exe_path': r"D:\ggxz\快期3\launcher.exe",
            # 'exe_path': r"D:\edge下载\快期3\launcher.exe",
            'username': config.get("customerId", ""),
            'password': config.get("tradePassword", ""),
            'investmentBankingName': config.get("investmentBankingName", ""),
            'PIN': config.get("PIN", ""),
            'distributorPhone': config.get("distributorPhone", ""),
            'timestamp': config.get("timestamp", "")
        }
        self.mod = "实盘"  # 模式 实盘/模拟
        # 当前所选券商的index，用于登录错误切换券商主席次席等
        self.login_broker_index = -1
        # 记录当前选择的券商名
        self.selected_broker = self.config['investmentBankingName']
        self.funds = None

        self.app_login = None  # 登录操作界面
        self.main_wnd = None # 主页面
        self.panel = None  # 交易面板
        self.action_panel = None  # 内嵌交易面板

        # 订单队列
        # 整体rpa操作执行队列
        # self.queue: queue.PriorityQueue = queue.PriorityQueue()
        self.high_priority_queue = queue.Queue()
        self.low_priority_queue = queue.Queue()
        self.worker_thread = None  # 订单处理线程，提前声明一个属性以便管理线程
        self.active = False
        self.condition = threading.Condition()

        self.MAX_RECORD_NUM = 0

    # 连接主程序，执行后续操作
    def connect(self) -> dict:
        # 0: 已登录 1: 登录成功 2: 当前处于非交易时间段，登录超时 3: 账号密码错误 10: 其他原因登录失败
        log_status = 10
        login_times = 0
        contract_list = []
        while login_times < 3:
            login_times += 1
            log.info(
                f"第{login_times}次登录客户端: <{self.config['username']}>[{self.selected_broker}]")
            log_status, contract_list = self.__login()
            if log_status == 1 or log_status == 0:
                log.info(
                    f"第{login_times}次登录成功: <{self.config['username']}>[{self.selected_broker}]")
                break
            log.error(
                f"第{login_times}次登录失败: <{self.config['username']}>[{self.selected_broker}]")
            if log_status == 2 or log_status == 3:
                break
            time.sleep(10)
        if log_status != 0 and log_status != 1:
            message = ""
            if log_status == 2:
                message = "关联失败，请在交易时段再次尝试"
            elif log_status == 3:
                message = "关联失败，请检查账号信息"
            elif log_status == 10:
                message = "关联失败，请联系客服"
            log.warning(
                f"RPA 登录失败 账户: <{self.config['username']}>[{self.selected_broker}] 原因：{message}")
            return {
                'success': False,
                'message': message,
                'log_status': log_status,
                'ib_list': contract_list
            }

        log.info(f"正在连接客户端: {self.config['exe_path']}")
        link_times = 0
        while link_times < 10:
            link_times += 1
            log.info(f"第{login_times}次连接客户端")
            self.app = pywinauto.Desktop(backend="uia").window(title_re=f".*{self.selected_broker}.*")
            if self.app.exists():
                break
            time.sleep(1)
        log.info(f"已连接到: {self.config['exe_path']}")
        self.main_wnd = self.app  # self.main_wnd
        # 窗口最大化
        if not self.main_wnd.was_maximized():
            self.main_wnd.maximize()

        if self.mod == "实盘":
            time.sleep(5)
            confirm_window = self.main_wnd.child_window(control_type="Window", title_re=f".*确认结算单.*")
            if confirm_window.exists():
                confirm_window.window(class_name="Button", title="确认").click_input()

        # 不知道是什么以防万一出现
        if self.mod == "实盘":
            time.sleep(5)
        over_window = self.main_wnd.child_window(control_type="Window", title_re=f".*监控中心.*")
        if over_window.exists():
            over_window.close()

        time.sleep(self.config["sleep_start"])

        try:
            jiaoyi_panel = self.main_wnd.child_window(title="交易面板", control_type="Pane")
            # 等待它真正可用
            jiaoyi_panel.wait("exists ready visible", timeout=5).set_focus()
            self.panel = jiaoyi_panel
        except timings.TimeoutError:
            print("第一次定位 '交易面板' 失败，尝试按 F11 并重试...")
            self.main_wnd.set_focus()
            send_keys("{F11}")
            time.sleep(1)  # 等待 UI 更新
            try:
                jiaoyi_panel = self.main_wnd.child_window(
                    title="交易面板", control_type="Pane", top_level_only=False
                )
                jiaoyi_panel.wait("exists ready visible", timeout=10).set_focus()
                print("重试后成功找到 '交易面板' 并设置焦点！")
                self.panel = jiaoyi_panel
            except timings.TimeoutError:
                print("重试后依然无法定位 '交易面板'，请检查控件属性。")
                # 输出父窗口控件树，帮助调试
                self.main_wnd.print_control_identifiers()
                raise

        self.panel.set_focus()
        # 在 workspace_spec 上继续调用 child_window() 来获取“内嵌下单板”选项卡的 WindowSpecification 对象
        self.action_panel = self.panel.child_window(title="内嵌下单板", control_type="Tab")
        self.__change_price_type_0()

        # 获取委托单列表高度
        grid_zone = self.panel.window(control_id=0x3EB, class_name='XTPReport')
        self.grid_height = grid_zone.rectangle().height() - 21
        # 一条记录的高度
        ONE_RECORD_HEIGHT = 26
        # 委托单一共可以显示记录的条数
        self.MAX_RECORD_NUM = self.grid_height // ONE_RECORD_HEIGHT
        log.info(f"委托单列表最多显示记录条数: {self.MAX_RECORD_NUM}")

        return {
            'success': True,
            'message': 'RPA 登陆成功',
            'log_status': log_status,
            'ib_list': contract_list
        }

    def start_worker(self):
        t = threading.Thread(target=self.handle_task_loop, daemon=True)
        t.start()
        self.worker_thread = t
        log.info("交易任务线程已启动")

    def __login(self):
        return_contract_list = []
        try:
            # 1. 判断是否已登录（根据券商名称）
            already_logged_in_spec = pywinauto.Desktop(backend='uia').window(
                title_re=f".*{self.config['investmentBankingName']}.*"
            )
            if already_logged_in_spec.exists(timeout=1):
                log.info("已登录，连接客户端进行后续操作")
                return 0, return_contract_list

            # 2. 判断是否存在“快期”窗口
            kq_spec = pywinauto.Desktop(backend='uia').window(title_re=".*快期.*")
            if kq_spec.exists(timeout=1):
                # 如果检测到窗口，直接使用它
                self.app_login = kq_spec
                log.info(f"已连接到 {self.config['exe_path']}")
            else:
                # 检查应用程序路径是否存在
                if not os.path.exists(self.config['exe_path']):
                    log.info(f"错误: 应用程序路径不存在: {self.config['exe_path']}")
                    return 10, return_contract_list  # 如果路径不存在，则返回登录失败

                # 启动应用程序
                log.info(f"启动应用程序: {self.config['exe_path']}")
                subprocess.Popen(self.config['exe_path'])  #

                # 等待应用程序启动
                log.info(f"等待应用程序启动 ({self.config['sleep_start']}秒)...")
                time.sleep(self.config['sleep_start'])  # 等待指定的时间，以确保应用程序启动

                # 使用Desktop获取所有顶级窗口
                desktop = pywinauto.Desktop(backend="uia")  # 使用UI自动化（uia）后台获取桌面

                # 查找登录窗口
                login_window = None
                all_windows = desktop.windows()  # 获取所有顶级窗口

                for w in all_windows:
                    try:
                        if "快期3" in w.window_text():  # 查找包含“登录”字样的窗口
                            login_window = w
                            break
                    except:
                        continue

            # self.app_login.set_focus()
            # time.sleep(self.config["sleep_operation"])

            login_window = self.app_login
            time.sleep(self.config["sleep_operation"])
            time.sleep(self.config["sleep_login"])

            # 4. 如果是“实盘”，选择期货券商
            if self.mod == '实盘':
                # 点击“打开”按钮以展开下拉框
                button_open = login_window.child_window(control_type='Button', title='打开', found_index=0)
                button_open.click_input()

                # 获取 List 控件对象
                contract_list_ctrl = login_window.child_window(class_name="ComboLBox",
                                                               control_type="List").wrapper_object()
                # 获取所有文本（有时返回的是嵌套列表，需要拍平）
                contract_list_texts = contract_list_ctrl.texts()
                flat_list = [item for sublist in contract_list_texts for item in sublist]
                return_contract_list = flat_list

                # 筛选出包含目标券商的条目
                exist_contract_list = [s for s in flat_list if self.config["investmentBankingName"] in s]
                if not exist_contract_list:
                    log.error(f"列表中未找到包含 '{self.config['investmentBankingName']}' 的项")
                    return 10, return_contract_list

                max_len = len(exist_contract_list)
                self.login_broker_index = (self.login_broker_index + 1) % max_len
                target_text = exist_contract_list[self.login_broker_index]

                try:
                    # 在 flat_list 中找目标文本的索引
                    target_index = flat_list.index(target_text)
                    target_item = contract_list_ctrl.get_item(target_index)
                    # 选中目标条目并按回车
                    target_item.select()
                    pywinauto.keyboard.send_keys('{ENTER}')
                    log.info("选择期货券商成功")
                except Exception as e:
                    log.error(f"选择期货券商错误: {e}")
            time.sleep(self.config["sleep_operation"])

            # 5. 输入用户名和密码
            funding_account = login_window.child_window(class_name="Edit", found_index=0)
            funding_account.click_input()
            time.sleep(0.5)
            pywinauto.keyboard.send_keys('^a{BACKSPACE}')  # 全选并删除
            pywinauto.keyboard.send_keys(self.config["username"])
            time.sleep(0.5)

            # 切换到密码输入框（Tab）
            pywinauto.keyboard.send_keys('{TAB}')
            time.sleep(0.5)
            pywinauto.keyboard.send_keys('^a{BACKSPACE}')
            pywinauto.keyboard.send_keys(self.config["password"])
            time.sleep(0.5)

            if self.config["PIN"]:
                pywinauto.keyboard.send_keys('{TAB}')
                time.sleep(0.5)
                pywinauto.keyboard.send_keys('^a{BACKSPACE}')
                pywinauto.keyboard.send_keys(self.config["PIN"])
                time.sleep(0.5)

            # 6. 查找并点击登录按钮
            buttons = login_window.descendants(control_type="Button")
            login_button = None
            for btn in buttons:
                try:
                    btn_text = btn.window_text()
                    if btn_text in ["登 录", "登录"]:
                        login_button = btn
                        break
                except:
                    continue
            # 如果没找到明确的登录按钮，则使用最后一个按钮以防万一
            if not login_button and buttons:
                login_button = buttons[-1]
            if login_button:
                login_button.click_input()
            time.sleep(self.config["sleep_operation"])

            # 7. 判断是否登录成功：等待登录窗口消失
            try:
                login_window.wait_not('exists', timeout=5)
                log.info("登录窗口消失，已登录成功")
                return 1, return_contract_list
            except:
                # 如果登录窗口依旧存在，可能出错或尚未完成
                pass

            time.sleep(self.config["sleep_login"])
            # 8. 检查错误信息 TODO 不知道是否完美
            err_ctrl = login_window.child_window(control_type="Text", title_re=".*失败.*")
            if err_ctrl.exists():
                text_content = err_ctrl.window_text()
                log.error(f"登录失败，错误信息: {text_content}")
                if "请求超时" in text_content:
                    return 2, return_contract_list
                elif "用户名" in text_content or "CTP" in text_content:
                    return 3, return_contract_list

            # 再等 2 秒后看看窗口是否依旧在
            time.sleep(2)
            if not login_window.exists():
                # 有时窗口消失得比较慢，第二次检查成功就代表登录成功
                return 1, return_contract_list

            # 如果依旧没成功，就超时
            return 10, return_contract_list

        except Exception as e:
            log.error(f"登录快期报错: {e}\n{traceback.format_exc()}")
            return 10, return_contract_list

    def trade(self, trade_request: TradeRequest):
        # 从封装对象中获取各个参数
        symbol = trade_request.symbol
        volume = trade_request.volume
        price = trade_request.price
        operation = trade_request.operation
        log.info("开始交易操作: symbol=%s, volume=%s, price=%s, operation=%s", symbol, volume, price, operation)

        # 获取输入控件
        symbol_window = self.action_panel.child_window(control_type="Edit", found_index=0)

        log.debug("设置焦点到 symbol 输入框")
        symbol_window.set_focus()
        time.sleep(self.config["sleep_start"])

        log.debug("点击 symbol 输入框")
        symbol_window.click_input()
        time.sleep(self.config["sleep_operation"])

        log.debug("清空 symbol 输入框")
        send_keys('{BACKSPACE}')
        log.info("输入 symbol: %s", symbol)
        send_keys(symbol)
        time.sleep(self.config["sleep_operation"])

        log.debug("按下 TAB 键切换到数量输入框") # 切换到数量
        send_keys('{TAB}')
        time.sleep(self.config["sleep_operation"])

        log.info("输入 volume: %s", volume)
        send_keys(str(volume))
        time.sleep(self.config["sleep_operation"])

        log.debug("按下 TAB 键切换到价格输入框")
        send_keys('{TAB}')
        time.sleep(self.config["sleep_operation"])

        log.debug("全选并清空价格输入框")
        send_keys('^a')
        time.sleep(self.config["sleep_operation"])
        send_keys('{BACKSPACE}')
        time.sleep(self.config["sleep_operation"])

        log.info("输入 price: %s", price)
        send_keys(str(price))
        time.sleep(self.config["sleep_operation"])

        # 根据 operation 参数选择对应的操作按钮
        log.info("选择操作按钮, operation index: %s", operation)
        operation_button = self.action_panel.child_window(control_type="Button", found_index=operation)
        operation_button.click_input()
        log.info("点击操作按钮完成，交易请求发送成功")


    def order_result(self) -> tuple[bool, str]:
        try:
            # 情况2：先判断“提示”窗口（只判断一次）
            dialog = self.main_wnd.child_window(title_re=".*提示.*", control_type="Window")
            if dialog.exists(timeout=1):
                log.info("当前交易已断线，不能执行此操作")
                dialog.child_window(control_type="Button", title_re=".*确定.*").click_input()
                return False, "当前交易已断线，不能执行此操作"  # 直接返回，不继续后续逻辑
        except Exception as e:
            log.warning("处理‘提示’窗口异常：%s", e)

        try:
            # 获取目标对话框
            dialog = self.main_wnd.child_window(control_type="Window", title_re=".*下单失败.*")
            if dialog.exists(timeout=1):
                log.info('下单失败')
                # 获取 Edit 控件
                doc = dialog.child_window(control_type="Edit")
                # 获取控件的 Value 属性
                value_text = doc.get_value()  # 获取 Value 值，这里是你需要的文本
                # 使用 log 输出获取的文本内容
                log.info("获取的文本内容：%s", value_text)
                # 使用正则表达式提取备注信息
                remark_match = re.search(r"备注：(.*)", value_text)
                # 输出匹配结果
                if remark_match:
                    log.info("提取的备注: %s", remark_match.group(1))
                else:
                    log.info("没有找到备注信息")
                # 点击“确定”按钮
                dialog.child_window(control_type="Button", title_re=".*确定.*").click_input()
                return False, f"下单失败: {remark_match.group(1) if remark_match else '无备注'}"
        except Exception as e:
            log.warning("处理‘下单失败’窗口异常：%s", e)

        try:
            # 情况3：再判断“注意”窗口（只判断一次）
            dialog = self.main_wnd.child_window(title_re=".*注意.*", control_type="Window")
            if dialog.exists(timeout=3):
                log.info("找到‘注意’窗口")
                dialog.child_window(control_type="Button", title_re=".*确定.*").click_input()
                return False, "操作被阻止：找到‘注意’窗口"
        except Exception as e:
            log.warning("处理‘注意’窗口异常：%s", e)

        try:
            # 判断是否下单成功
            log.info("判断是否下单成功")
            order_result = self.is_establish()
            if order_result:
                # 再进行一次判断，防止误判
                dialog = self.main_wnd.child_window(title_re=".*成交通知.*", control_type="Window")
                if dialog.wait("exists"):
                    doc_ctrl = dialog.child_window(control_type="Edit", auto_id="1074").wrapper_object()
                    text_content = doc_ctrl.get_value()
                    log.info("成交通知窗口的文本内容：")
                    log.info(text_content)
                    # 匹配成交量和成交价
                    pattern_volume = r"成交量：\s*([0-9.]+)"
                    pattern_price = r"成交价：\s*([0-9.]+)"

                    match_volume = re.search(pattern_volume, text_content)
                    match_price = re.search(pattern_price, text_content)

                    if match_volume and match_price:
                        volume = match_volume.group(1)
                        price = match_price.group(1)
                        print("成交量:", volume)
                        print("成交价:", price)
                        self.funds = self.funds - (volume * price)
                    else:
                        print("未找到成交量或成交价的数据")

                    dialog.child_window(control_type="Button", title_re=".*确定.*").click_input()
                    return True, text_content
            else:
                return False, "下单失败"
        except Exception as e:
            log.warning("判断下单是否成功失败")

        try:
            # 情况1：等待“成交通知”窗口最长120秒
            log.info("等待‘成交通知’窗口，最多120秒")
            dialog = self.main_wnd.child_window(title_re=".*成交通知.*", control_type="Window")
            if dialog.wait("exists", timeout=110):
                doc_ctrl = dialog.child_window(control_type="Edit", auto_id="1074").wrapper_object()
                text_content = doc_ctrl.get_value()
                log.info("成交通知窗口的文本内容：")
                log.info(text_content)
                # 匹配成交量和成交价
                pattern_volume = r"成交量：\s*([0-9.]+)"
                pattern_price = r"成交价：\s*([0-9.]+)"

                match_volume = re.search(pattern_volume, text_content)
                match_price = re.search(pattern_price, text_content)

                if match_volume and match_price:
                    volume = match_volume.group(1)
                    price = match_price.group(1)
                    print("成交量:", volume)
                    print("成交价:", price)
                    self.funds = self.funds - (volume * price)
                else:
                    print("未找到成交量或成交价的数据")
                dialog.child_window(control_type="Button", title_re=".*确定.*").click_input()
                return True, text_content
            else:
                raise Exception("120秒内未检测到‘成交通知’窗口")
        except Exception as e:
            log.error("处理‘成交通知’异常：%s", e)
            log.info("进入自定义失败处理逻辑:撤单")
            self.cancel_order()
            return False, f"成交通知未在120秒内出现，已撤单"

    # 选择市价
    def __change_price_type_0(self):
        price_type = self.action_panel.window(
            control_id=0xce7, class_name="Edit")
        price_type.click_input()
        time.sleep(0.1)
        price_type.click_input(coords=(100, 40))


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

    from datetime import datetime, timedelta, time as dt_time

    def handle_task_loop(self):
        log.info("任务处理进程启动")
        # 定义允许处理任务的时间段列表
        allowed_periods = [
            (dt_time(8, 45), dt_time(11, 35)),  # 上午段：08:45 ~ 11:35
            (dt_time(13, 0), dt_time(15, 5)),  # 下午段：13:00 ~ 15:05
            (dt_time(20, 45), dt_time(23, 5))  # 夜盘段：20:45 ~ 23:05
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
                log.info("执行资金查询（get_funds）接口～")
                self.get_funds()

                next_funds_call = datetime.now() + timedelta(minutes=15)

            task = self.get_next_task()  # 自动阻塞等待任务
            if not task:
                continue  # 如果没有任务则跳过本次循环

            max_retries = 3
            for attempt in range(1, max_retries + 1):
                try:
                    log.info(f"执行交易任务 {task}，第 {attempt} 次尝试")
                    if self.funds < task.volume * task.price:
                        return False
                    self.trade(task)  # 执行交易任务
                    result = self.order_result()  # 等待订单结果

                    # 检查订单结果是否不为空
                    if result and result[0]:
                        log.info(f"交易任务 {task} 成功完成")
                        break  # 成功处理，跳出重试循环
                    elif result and not result[0]:
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

    def restart(self) -> dict[str, str | int | dict[str, str | Any]] | None:
        result = self.connect()

        if isinstance(result, dict) and result.get("success"):
            log.info("定时重启成功")
            return {
                'type': '定时服务',
                'status': 0,
                'info': {'账号': self.config["username"], '用户': self.config["user_id"]}
            }
        else:
            log.warning(f"定时重启失败，返回结果: {result}")
            return {
                'type': '定时重启服务',
                'status': 1,
                'info': '重启失败'
            }

    def close(self) -> dict[str, str | int | dict[str, str | Any]] | None:
        try:
            # 窗口最大化
            if not self.main_wnd.was_maximized():
                self.main_wnd.maximize()
            self.main_wnd.close()
            close_window = self.main_wnd.child_window(control_type="Window", title_re=".*注意.*")
            confirm_button = close_window.child_window(control_type="Button", title_re='.*是.*')
            confirm_button.click_input()
            return {'type': '定时服务', 'status': 1,
                    'info': {'账号': self.config["username"], '用户': self.config["user_id"]}}
        except Exception as e:
            log.error('关闭程序时出错')

    def start_schedule(self):
        # 定义定时任务的回调函数
        # 开盘前重启操作，适应市场交易时间
        # 设置东八区时区
        tz = pytz.timezone('Asia/Shanghai')

        schedule.every().day.at('08:45').do(self.restart).tag(tz)  # 开盘前15分钟重启
        schedule.every().day.at('13:00').do(self.restart).tag(tz)  # 下午开盘前30分钟重启
        schedule.every().day.at('20:45').do(self.restart).tag(tz)  # 夜盘开盘前15分钟重启

        schedule.every().day.at("11:35").do(self.close).tag(tz)  # 上午休盘后关闭
        schedule.every().day.at("15:05").do(self.close).tag(tz)  # 下午休盘后关闭
        schedule.every().day.at("23:05").do(self.close).tag(tz)  # 夜盘休盘后关闭

        # # 设置定时任务，使用动态计算的时间（测试使用）
        # # 计算新的时间（分别加上2分钟和3分钟）
        # restart_time = datetime.now() + timedelta(minutes=2)  # 2分钟后重启
        # close_time = datetime.now() + timedelta(minutes=1)  # 1分钟后关闭
        # schedule.every().day.at(restart_time.strftime('%H:%M')).do(self.restart).tag(tz)
        # schedule.every().day.at(close_time.strftime('%H:%M')).do(self.close).tag(tz)
        #
        # print(f"重启任务设置为 {restart_time.strftime('%H:%M')}")
        # print(f"关闭任务设置为 {close_time.strftime('%H:%M')}")

        # 启动一个单独线程运行定时调度循环
        def run_schedule():
            while True:
                schedule.run_pending()
                time.sleep(1)

        t = threading.Thread(target=run_schedule, daemon=True)
        t.start()
        log.info("定时调度任务已启动")

    def cancel_order(self):
        # 定位包含委托单的 Tab 控件
        order_table = self.panel.child_window(control_type='Tab', title_re=".*委托单.*")

        xtpBarRight_pane = order_table.child_window(control_type='Pane', title_re=".*xtpBarRight.*")
        xtpBarRight_pane.child_window(control_type='Button', title_re=".*撤单.*").click_input()
        cancel_order_window = self.main_wnd.child_window(control_type="Window", title_re=".*确认撤单.*")
        cancel_order_window.child_window(control_type="Button", title_re=".*确定.*").click_input()
        cancel_order_success_window = self.main_wnd.child_window(control_type="Window", title_re=".*撤单成功.*")
        # cancel_order_success_window.print_control_identifiers()
        cancel_order_success_doc = cancel_order_success_window.child_window(control_type='Edit').wrapper_object()
        cancel_order_success_doc_text = cancel_order_success_doc.get_value()
        log.info("撤单内容：" + cancel_order_success_doc_text)
        cancel_order_success_window.child_window(control_type="Button", title_re=".*确定.*").click_input()

    def is_establish(self):
        position = self.panel.child_window(control_type="TabItem", title_re=".*委托单.*", found_index=0)
        position.click_input(button='left')
        history_order = self.panel.child_window(control_type="Tab", title_re=".*委托单.*")
        element = history_order.child_window(control_type="Header", title_re=".*详细状态.*")
        rect = element.rectangle()
        l, t, r, b = get_bottom_rect(rect.left, rect.top, rect.right, rect.bottom)
        # 点击第一条交易信息
        pywinauto.mouse.click(button='left', coords=(round((l + r) / 2), round((t + b) / 2)))
        report_table = self.panel.child_window(control_type="Table", title="Report")
        # report_table.print_control_identifiers()
        # 做到一个刷新页面的效果
        self.main_wnd.minimize()
        self.main_wnd.maximize()
        time.sleep(3)
        row = report_table.child_window(title_re=".*Row.*", control_type="Custom", found_index=0)
        res = row.child_window(control_type="DataItem", found_index=10).wrapper_object()
        name = res.window_text()
        log.info(f"控件的 name 是: {name}")
        print(type(res.window_text()))
        return "不成交" == name

    def export_csv(self):
        position = self.panel.child_window(control_type="TabItem", title_re=".*持仓.*", found_index=0)
        position.click_input(button='right')
        self.main_wnd.child_window(control_type="MenuItem", title_re=".*csv.*").click_input()
        export = self.main_wnd.child_window(control_type="Window", title_re="导出表格")
        # export.print_control_identifiers()
        # 获取当前用户主目录
        home_dir = os.path.expanduser('~')
        # 根据系统决定使用 'Desktop' 或 '桌面' 文件夹名称
        desktop_folder = 'Desktop'  # 如果是中文系统则可能为 '桌面'
        desktop_path = os.path.join(home_dir, desktop_folder)

        # 拼接出桌面上的 futures 文件夹路径
        futures_folder = os.path.join(desktop_path, 'futures')

        # 检查并创建
        if not os.path.exists(futures_folder):
            os.makedirs(futures_folder)
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

        send_keys(futures_folder)
        time.sleep(1)
        send_keys('{ENTER}')

        export.child_window(control_type="Button", title_re=".*保存.*").click_input()
        confirm = export.child_window(control_type="Button", title_re=".*是.*")
        try:
            # 等待控件存在（最多 5 秒钟）
            confirm.wait('exists', timeout=2)
            confirm.click_input()
        except timings.TimeoutError:
            pass
        (self.main_wnd.child_window(control_type="Window", title_re=".*注意.*")
         .child_window(control_type="Button", title_re=".*确定.*").click_input())

    def analysis_csv(self):
        # 获取当前用户主目录
        home_dir = os.path.expanduser('~')
        # 根据系统决定使用 'Desktop' 或 '桌面' 文件夹名称
        desktop_folder = 'Desktop'  # 如果是中文系统则可能为 '桌面'
        desktop_path = os.path.join(home_dir, desktop_folder)

        # 拼接出桌面上的 futures 文件夹路径
        futures_folder = os.path.join(desktop_path, 'futures')

        # 检查并创建
        if not os.path.exists(futures_folder):
            os.makedirs(futures_folder)

        csv_files = [file for file in os.listdir(futures_folder) if file.lower().endswith('.csv')]


        if csv_files:
            log.info("在 'futures' 文件夹中找到以下 CSV 文件:")
            for csv_file in csv_files:
                print(" -", csv_file)

            # 默认解析找到的第一个 CSV 文件
            selected_csv = csv_files[0]
            csv_file_path = os.path.join(futures_folder, selected_csv)
            print("\n开始解析 CSV 文件：", selected_csv)

            def parse_float(s: str) -> float:
                """
                尝试将字符串转换为 float，
                去掉末尾的百分号%，若为空或无效，则返回0.0。
                """
                if not s or s.strip() == '-' or s.strip() == '':
                    return 0.0
                s = s.replace('%', '').strip()
                try:
                    return float(s)
                except ValueError:
                    return 0.0

            def parse_int(s: str) -> int:
                """
                尝试将字符串转换为 int，
                若为空或无效，则返回0。
                """
                if not s or s.strip() == '-' or s.strip() == '':
                    return 0
                try:
                    return int(s)
                except ValueError:
                    return 0

            try:
                # 尝试用 GBK 编码读取 CSV 文件
                with open(csv_file_path, newline='', encoding="gbk") as csvfile:
                    reader = csv.reader(csvfile)
                    rows = list(reader)

                if len(rows) < 2:
                    raise ValueError("CSV 文件没有有效数据")

                    # 第一行是表头，后面每行为数据
                data_rows = rows[1:]
                history_orders = []

                for row in data_rows:
                    # 如果遇到合计行或者已经收集了5条，就停止处理
                    if row and row[0].startswith("合计"):
                        break
                    if len(history_orders) >= 5:
                        break

                    # 逐列解析并封装到 HistoryOrder
                    ho = HistoryOrder(
                        contract=row[0].strip(),
                        side=row[1].strip(),
                        total_position=parse_int(row[2]),
                        open_price=parse_float(row[3]),
                        floating_pnl=parse_float(row[4]),
                        floating_pnl_ratio=parse_float(row[5]),
                        quoted_pnl=parse_float(row[6]),
                        actual_margin=parse_float(row[7]),
                        capital_ratio=parse_float(row[8]),
                        manual_stop_loss=row[9].strip(),
                        manual_take_profit=row[10].strip(),
                        stop_loss_volume=row[11].strip(),
                        auto_stop_loss=row[12].strip(),
                        auto_take_profit=row[13].strip(),
                        position_value=parse_float(row[14]),
                        position_type=row[15].strip(),
                        delta=parse_float(row[16]),
                        gamma=parse_float(row[17]),
                        theta=parse_float(row[18]),
                        vega=parse_float(row[19]),
                        rho=parse_float(row[20]),
                        time_value=parse_float(row[21]),
                        expiration_date=row[22].strip() if len(row) > 22 else ""
                    )
                    history_orders.append(ho)

                return history_orders

            except UnicodeDecodeError as e:
                raise Exception(
                    "CSV 文件解析失败：请检查文件编码，尝试使用其他编码（例如 'gbk' 或 'latin1'）。\n错误详情：{}".format(e))
        else:
            raise FileNotFoundError("没有找到 CSV 文件！")

    def operation_positions(self, index, operation):
        position = self.panel.child_window(control_type="TabItem", title_re=".*持仓.*", found_index=0)
        position.click_input(button='left')
        element = self.panel.child_window(control_type="Header", title_re=".*持仓合约.*")
        rect = element.rectangle()
        l, t, r, b = rect.left, rect.top, rect.right, rect.bottom
        if (index and operation != 4):
            for i in range(index):
                l, t, r, b = get_bottom_rect(l, t, r, b)
            pywinauto.mouse.click(button='left', coords=(round((l + r) / 2), round((t + b) / 2)))
            report_table = self.panel.child_window(control_type="Table", title="Report")
            # report_table.print_control_identifiers()
            # 做到一个刷新页面的效果
            self.main_wnd.minimize()
            self.main_wnd.maximize()
            time.sleep(3)
        operations = {
            1: self.panel.child_window(control_type="Button", title_re=".*市价平仓.*"),
            2: self.panel.child_window(control_type="Button", title_re=".*市价反手.*"),
            3: self.panel.child_window(control_type="Button", title_re=".*行权.*"),
            4: self.panel.child_window(control_type="Button", title_re=".*全部清仓.*")
        }

        button = operations.get(operation)

        if button:
            button.click_input()
            context = self.main_wnd.child_window(control_type="Window", title_re=".*确认.*")
            res_element = context.child_window(control_type="Text")
            # 转为 wrapper 对象
            text_wrapper = res_element.wrapper_object()

            # 获取 Name 属性和 window_text() 返回的内容
            name_text = text_wrapper.element_info.name
            window_text = text_wrapper.window_text()

            print("Name属性:", name_text)
            print("window_text 返回内容:", window_text)

            # 尝试获取 legacy_properties，如果不存在则捕获异常
            try:
                legacy_props = text_wrapper.element_info.legacy_properties
                legacy_value = legacy_props.get("Value", "") if legacy_props else ""
                print("Legacy Value属性:", legacy_value)
            except AttributeError:
                print("Legacy Value属性: 不可用")
            confirm_button = context.child_window(control_type="Button", title_re=".*确定.*")
            confirm_button.click_input()

            back_window = self.main_wnd.child_window(control_type="Window", title_re=".*下单.*")
            doc = back_window.child_window(control_type="Edit")
            value_text = doc.get_value()
            log.info(value_text)
            msg = re.search(r"备注：(.+)", value_text)
            remark = msg.group(1)
            back_window.child_window(control_type="Button", title_re=".*确定.*").click_input()
            if "拒绝" in remark:
                return False, remark
            else :
                return True, remark


    def test(self):
        pass

    def get_funds(self):
        x, y = 1705, 17
        # TODO 使用鼠标左键点击指定坐标,可能在其他的地方无法实现
        mouse.click(button='left', coords=(x, y))
        self.main_wnd.child_window(control_type="MenuItem", title_re=".*查询资金.*", found_index=0).click_input()
        time.sleep(3)
        funds_window = self.main_wnd.child_window(control_type="", title_re=".*期货资金账户详情.*")
        funds_msg = funds_window.child_window(control_type="Edit")
        edit_ctrl = funds_msg.wrapper_object()

        # 尝试获取文本内容
        text_content = edit_ctrl.get_value()
        print(text_content)
        funds_window.close()
        # 定义正则表达式模式，匹配“可用资金：”后面的数值
        pattern = r"可用资金：\s*([-0-9.]+)"
        match = re.search(pattern, text_content)

        if match:
            self.funds = match.group(1)
            print("可用资金:", self.funds)
        else:
            print("未找到可用资金的数据")


if __name__ == "__main__":
    sample_config = {
        "customerId": "228855",
        "tradePassword": "koumin917#",
        "investmentBankingName": "simnow",
        "PIN": "",
        "distributorPhone": "user001",
        "timestamp": "ib001"
    }
    rpa_instance = RpaOperator(sample_config)
    rpa_instance.connect()
    # rpa_instance.start_worker()
    rpa_instance.get_funds()