

# 内部使用-存储交易后的数据与交易端获取的dict
class TradeTempData:
    """
    用于临时保存交易数据，便于后续对比、撤单重试等操作。

    Attributes:
        order_obj: GatewayOrderData 对象，当前订单数据。
        first_order_obj: GatewayOrderData 对象，首次下单时的订单数据（用于撤单前对比）。
        first_origin_dict: dict，首次下单时的原始数据字典。
        last_check_order_obj: GatewayOrderData 对象，上次查询时保存的订单数据。
        last_check_origin_dict: dict，上次查询时保存的原始数据字典。
        commit_volume: int，下单手数，初次下单以及撤单重试时更新。
    """
    def __init__(self, order_obj: GatewayOrderData, first_order_obj: GatewayOrderData = None, first_origin_dict: dict = {},
                 last_check_order_obj: GatewayOrderData = None, last_check_origin_dict: dict = {}, commit_volume: int = 0) -> None:
        # 当前订单数据对象
        self.order_obj: GatewayOrderData = order_obj
        # 撤单前保存的第一个订单对象（首次下单时的订单数据）
        self.first_order_obj: GatewayOrderData = first_order_obj
        # 撤单前保存的原始数据字典
        self.first_origin_dict: dict = first_origin_dict
        # 上次查询保存的订单数据对象
        self.last_check_order_obj: GatewayOrderData = last_check_order_obj
        # 上次查询保存的原始数据字典
        self.last_check_origin_dict: dict = last_check_origin_dict
        # 下单手数，初次下单和撤单重试时可能会更新
        self.commit_volume = commit_volume


# 内部使用-下单信息类
class OrderInfoItem:
    """
    存储下单操作相关的信息，用于记录下单时的实际参数和状态。

    Attributes:
        oparams: OperateParam 对象，下单操作的参数（如合约、买卖、价格等）。
        realAmount: int，实际下单手数。
        click_time: int，点击下单按钮的时间戳或计数，用于记录操作时间。
        real_price_type: int，实际使用的价格类型（比如市价、限价等）。
        trade_result: dict，下单操作返回的结果数据。
        retry: bool，是否允许重试下单操作（默认为 True）。
        revokeRetry: bool，是否允许撤单重试（默认为 True）。
    """
    def __init__(self, order_params: OperateParam, realAmount: int, click_time: int, real_price_type: int, trade_result: dict, retry: bool = True, revokeRetry: bool = True):
        # 下单操作参数
        self.order_params = order_params
        # 实际下单手数
        self.realAmount = realAmount
        # 点击下单的时间或次数
        self.click_time = click_time
        # 实际使用的价格类型
        self.real_price_type = real_price_type
        # 下单后返回的结果数据（通常是字典格式）
        self.trade_result = trade_result
        # 是否允许下单重试
        self.retry = retry
        # 是否允许撤单重试
        self.revokeRetry = revokeRetry


# 内部使用-订单信息类
class OrderItem:
    """
    存储订单信息，用于跟踪订单状态和执行相关查询。

    Attributes:
        oparams: OperateParam 对象，下单操作的参数。
        entrust_no: str，订单委托号，用于唯一标识订单。
        check_count: int，订单状态查询次数。
        check_time: int，订单状态最后一次查询的时间。
        order_time: int，下单的时间。
        realAmount: int，实际下单手数。
        real_price_type: int，实际下单使用的价格类型。
        retry: bool，是否允许重试下单操作（默认为 True）。
        revokeRetry: bool，是否允许撤单重试（默认为 True）。
    """
    def __init__(self, order_params: OperateParam, entrust_no: str, check_count: int, check_time: int, order_time: int, realAmount: int, real_price_type: int, retry: bool = True, revokeRetry: bool = True):
        # 下单操作参数
        self.order_params = order_params
        # 订单委托号
        self.entrust_no = entrust_no
        # 订单查询次数
        self.check_count = check_count
        # 最后一次查询订单状态的时间
        self.check_time = check_time
        # 下单时间
        self.order_time = order_time
        # 实际下单手数
        self.realAmount = realAmount
        # 实际使用的价格类型
        self.real_price_type = real_price_type
        # 是否允许下单重试
        self.retry = retry
        # 是否允许撤单重试
        self.revokeRetry = revokeRetry


# 内部使用-待查询数量的 close 类
class CheckCloseItem:
    """
    存储待查询平仓相关信息，用于后续的平仓操作和查询。

    Attributes:
        order_params: OperateParam 对象，下单操作的参数。
        click_time: int，点击查询平仓的时间戳或计数。
        result: dict，查询操作返回的结果数据。
        real_price_type: int，实际使用的价格类型。
        retry: bool，是否允许重试查询操作（默认为 True）。
        revokeRetry: bool，是否允许撤单重试（默认为 True）。
    """
    def __init__(self, order_params: OperateParam, click_time: int, result: dict, real_price_type: int, retry: bool = True, revokeRetry: bool = True):
        # 下单操作参数
        self.order_params = order_params
        # 点击查询平仓的时间或计数
        self.click_time = click_time
        # 查询返回的结果数据
        self.result = result
        # 实际使用的价格类型
        self.real_price_type = real_price_type
        # 是否允许查询重试
        self.retry = retry
        # 是否允许撤单重试
        self.revokeRetry = revokeRetry
