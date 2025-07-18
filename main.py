import os
import time
from dataclasses import asdict

from flask import Flask, request, jsonify

from pojo.order import TradeRequest, Order
import logging

from rpa.rpa_forward import RpaOperator
from rpa.rpa_operate import RPAOperate

app = Flask(__name__)
rpa_forward = None
rpa_operation = None

# 创建日志文件夹及日志
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

def make_response(code: int, message: str, data=None):
    """统一返回数据格式的助手函数"""
    return jsonify({
        'code': code,
        'message': message,
        'data': data
    }), code

@app.route('/futures/connect', methods=['POST'])
def start_rpa_api():
    """ 启动任务 """
    try:
        data = request.get_json()
        required_fields = [
            "customerId",
            "tradePassword",
            "investmentBankingName",
            "PIN",
            "distributorPhone",
            "timestamp"
        ]
        missing_fields = []
        # 检查必需参数是否存在，对于 PIN 字段只检查是否存在，不管值是否为空
        for field in required_fields:
            if field not in data:
                missing_fields.append(field)
            elif field != "PIN" and not data[field]:
                missing_fields.append(field)
        if missing_fields:
            return make_response(400, f"缺少必要参数或参数为空: {', '.join(missing_fields)}", None)

        sample_config = {
            "customerId": data["customerId"],
            "tradePassword": data["tradePassword"],
            "investmentBankingName": data["investmentBankingName"],
            "PIN": data["PIN"],
            "distributorPhone": data["distributorPhone"],
            "timestamp": data["timestamp"]
        }
        global rpa_forward
        from rpa.rpa_forward import RpaOperator
        rpa_forward = RpaOperator(sample_config)
        rpa_forward.connect()
        rpa_forward.start_worker()

        return make_response(200, 'RPA已启动', None)
    except Exception as e:
        log.error(f"启动RPA接口出错: {e}")
        return make_response(500, str(e), None)

@app.route('/futures/add_task', methods=['POST'])
def add_task_api():
    """ 添加任务 """
    if rpa_forward is None:
        return make_response(500, "未启动rpa", None)
    try:
        data = request.get_json()
        trade_request = TradeRequest(
            symbol=data['symbol'],
            volume=data['volume'],
            price=data['price'],
            operation=data['operation']
        )
        high_priority = data.get('high_priority', False)
        res = rpa_forward.add_task(trade_request, high_priority)
        if res:
            return make_response(200, '任务已添加', None)
        else:
            return make_response(500, "添加任务失败,请查看日志")
    except Exception as e:
        log.error(f"添加任务接口出错: {e}")
        return make_response(500, str(e), None)

@app.route('/futures/add_tasks', methods=['POST'])
def add_tasks_api():
    """ 添加批量任务 """
    if rpa_forward is None:
        return make_response(500, "未启动rpa", None)
    try:
        data = request.get_json()
        tasks_data = data.get('tasks', [])
        if not tasks_data:
            return make_response(400, '请求中缺少任务数据', None)

        high_priority = data.get('high_priority', False)
        trade_requests = []
        for task in tasks_data:
            trade_request = TradeRequest(
                symbol=task['symbol'],
                volume=task['volume'],
                price=task['price'],
                operation=task['operation']
            )
            trade_requests.append(trade_request)

        # 批量添加任务到队列
        rpa_forward.add_tasks(trade_requests, high_priority)
        return make_response(200, f'成功添加了 {len(trade_requests)} 个任务', None)
    except Exception as e:
        log.error(f"添加批量任务接口出错: {e}")
        return make_response(500, str(e), None)

@app.route('/futures/orders', methods=['GET'])
def get_history_orders():
    """ 获取所有订单数据 """
    if rpa_forward is None:
        return make_response(500, "未启动rpa", None)
    try:
        rpa_forward.export_csv()
        orders = rpa_forward.analysis_csv()
        # 将 dataclass 对象转换为字典，便于 jsonify 序列化
        orders_data = [asdict(order) for order in orders]
        return make_response(200, '获取订单成功', orders_data)
    except Exception as e:
        log.error(f"CSV 接口出错: {e}")
        return make_response(500, str(e), None)

@app.route('/futures/order', methods=['GET'])
def get_history_one_order():
    """ 获取某一条订单数据 """
    if rpa_forward is None:
        return make_response(500, "未启动rpa", None)
    try:
        # 导出 CSV 并解析数据
        rpa_forward.export_csv()
        orders = rpa_forward.analysis_csv()
        orders_data = [asdict(order) for order in orders]

        # 获取查询参数
        contract = request.args.get('contract')
        side = request.args.get('side')

        # 如果存在 contract 或 side 参数，则进行过滤
        if contract or side:
            filtered_orders = []
            for order in orders_data:
                if contract and order.get('contract') != contract:
                    continue
                if side and order.get('side') != side:
                    continue
                filtered_orders.append(order)
            orders_data = filtered_orders

        return make_response(200, '获取订单成功', orders_data)
    except Exception as e:
        log.error(f"CSV 接口出错: {e}")
        return make_response(500, str(e), None)

@app.route('/futures/operation', methods=['POST'])
def operation_positions():
    """ 操作持仓（POST 请求，参数通过请求体传入）"""
    if rpa_forward is None:
        return make_response(500, "未启动rpa", None)
    try:
        # 导出 CSV 并解析数据
        rpa_forward.export_csv()
        orders = rpa_forward.analysis_csv()
        time.sleep(2)
        orders_data = [asdict(order) for order in orders]

        # 从请求体中获取参数
        data = request.get_json()
        contract = data.get('contract')
        side = data.get('side')
        operation = int(data.get('operation'))

        index = None
        for i, record in enumerate(orders_data):
            if record.get("contract") == contract and record.get("side") == side:
                print(f"找到匹配项，它是第 {i + 1} 条数据")
                index = i + 1  # RPA可能需要 1-based index
                break

        if index is None:
            raise ValueError("没有找到符合条件的数据")

        res, msg = rpa_forward.operation_positions(index, operation)
        if res:
            return make_response(200, msg, None)
        else:
            return make_response(200, msg, None)
    except Exception as e:
        log.error(f"操作持仓接口出错: {e}")
        return make_response(500, str(e), None)


@app.route('/stock/cancel_order', methods=['POST'])
def cancel_order():
    """
    撤单接口
    POST /api/cancel_order
    请求 JSON:
    {
      "symbol": "...",
      "operation": "...",
      "contract_number": "..."
      // 更多字段...
    }
    """
    data = request.get_json()
    if not data:
        return make_response(400, "请求体必须是 JSON")

    try:
        order = Order(
            symbol=data.get('symbol'),
            operation=data.get('operation'),
            contract_number=data.get('contract_number'),
            # 填充其他必需字段...
        )
    except Exception as e:
        log.error(f"❌ 构造 Order 对象失败: {e}")
        return make_response(400, f"参数错误：{e}")

    try:
        success: bool = rpa_operation.cancel_task(order)
        if success:
            return make_response(200, "撤单成功")
        else:
            return make_response(500, "撤单失败，请检查日志获取更多信息")
    except Exception as ex:
        log.exception("❌ 撤单过程出现异常")
        return make_response(500, f"撤单异常：{ex}")

@app.route('/stock/add_task', methods=['POST'])
def add_task_stock():
    """ 添加任务 """
    if rpa_operation is None:
        return make_response(500, "未启动rpa", None)
    try:
        data = request.get_json()
        trade_request = TradeRequest(
            symbol=data['symbol'],
            volume=data['volume'],
            price=data['price'],
            operation=data['operation']
        )
        high_priority = data.get('high_priority', False)
        res = rpa_operation.add_task(trade_request, high_priority)
        if res:
            return make_response(200, '任务已添加', None)
        else:
            return make_response(500,"添加任务失败,请查看日志")
    except Exception as e:
        log.error(f"添加任务接口出错: {e}")
        return make_response(500, str(e), None)

@app.route('/stock/add_tasks', methods=['POST'])
def add_tasks_stock():
    """ 添加批量任务 """
    if rpa_operation is None:
        return make_response(500, "未启动rpa", None)
    try:
        data = request.get_json()
        tasks_data = data.get('tasks', [])
        if not tasks_data:
            return make_response(400, '请求中缺少任务数据', None)
        high_priority = data.get('high_priority', False)
        trade_requests = []
        for task in tasks_data:
            trade_request = TradeRequest(
                symbol=task['symbol'],
                volume=task['volume'],
                price=task['price'],
                operation=task['operation']
            )
            trade_requests.append(trade_request)
        # 批量添加任务到队列
        rpa_operation.add_tasks(trade_requests, high_priority)
        return make_response(200, f'成功添加了 {len(trade_requests)} 个任务', None)
    except Exception as e:
        log.error(f"添加批量任务接口出错: {e}")
        return make_response(500, str(e), None)

@app.route('/stock/pending_orders', methods=['GET'])
def pending_orders():
    """ 获取历史委托单 """
    if rpa_operation is None:
        return make_response(500, 'RPA 实例未启动', None)
    try:
        data_list = rpa_operation.pending_orders()
        # 假设 data_list 是个字典列表，可以直接 jsonify
        return make_response(200, '历史委托获取成功', data_list)
    except Exception as e:
        log.error(f"pending_orders 接口异常: {e}")
        return make_response(500, str(e), None)

@app.route('/stock/connect', methods=['POST'])
def connect_endpoint():
    """
    接收 JSON 格式的 config，例如：
    {
      "trade_account": "029001059286",
      "trade_password": "123457",
      "brokerage": "东吴证券",
      "opening_area": "",
      "sales_department": "",
      "business": ""
    }
    """
    global rpa_operation
    config = request.get_json(force=True)
    required_keys = [
        "trade_account", "trade_password", "brokerage",
        "opening_area", "sales_department", "business"
    ]
    missing = [k for k in required_keys if k not in config]
    if missing:
        return make_response(400, f"缺少参数：{', '.join(missing)}")

    try:
        # 用前端传来的 config 初始化并调用连接
        rpa_operation = RPAOperate(config)
        result = rpa_operation.connect()
        rpa_operation.start_worker()
        if result.get('success'):
            return make_response(200, result['message'])
        else:
            return make_response(500, result['message'])
    except Exception as e:
        return make_response(500, f"内部错误：{e}")


@app.route('/stock/history_orders', methods=['GET'])
def history_orders_api():
    """ 获取历史委托单 """
    if rpa_operation is None:
        return make_response(500, 'RPA 实例未启动', None)
    try:
        data_list = rpa_operation.history_orders()
        return make_response(200, '历史委托获取成功', data_list)
    except Exception as e:
        log.error(f"history_orders 接口异常: {e}")
        return make_response(500, str(e), None)

if __name__ == "__main__":
    # sample_config = {
    #     "customerId": "228855",
    #     "tradePassword": "koumin917#",
    #     "investmentBankingName": "simnow",
    #     "PIN": "",
    #     "distributorPhone": "user001",
    #     "timestamp": "ib001"
    # }
    # rpa_instance = RpaOperator(sample_config)
    # rpa_instance.connect()
    # rpa_instance.start_worker()
    print(int("1") == 1)
    # 启动 HTTP 服务
    app.run(host='0.0.0.0', port=5000)
