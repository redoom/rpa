from dataclasses import dataclass


@dataclass
class TradeRequest:
    """
    交易请求参数数据类，用于封装交易所需参数。
    参数说明：
      symbol: 交易标的代码
      volume: 交易数量
      price: 交易价格
      operation: 操作类型:期货:（0:买开/买多, 1:卖开/买空/锁仓, 2:买平, 3:卖平/平多）
                        股票: （0:买入股票, 1:卖出股票）
    """
    symbol: str
    volume: int
    price: float
    operation: int

@dataclass
class Order:
    """
    股票撤单使用
    参数说明：
        symbol:交易标的代码
        operation:方向（买入/卖出）
        contract_number:合同编号
    """
    symbol: str
    operation: str
    contract_number: str

@dataclass
class HistoryOrder:
    """
         持仓合约     -> contract
         买卖         -> side
         总仓         -> total_position
         开仓均价     -> open_price
         浮动盈亏     -> floating_pnl
         浮盈比例     -> floating_pnl_ratio
         对价盈亏     -> quoted_pnl   (此处翻译见备注)
         实收保证金   -> actual_margin
         资金占比     -> capital_ratio
         手工止损     -> manual_stop_loss
         手工止盈     -> manual_take_profit
         止损手数     -> stop_loss_volume
         自动止损     -> auto_stop_loss
         自动止盈     -> auto_take_profit
         持仓市值     -> position_value
         虚实         -> position_type (具体含义可进一步明确)
         持仓Delta    -> delta
         持仓Gamma    -> gamma
         持仓Theta    -> theta
         持仓Vega     -> vega
         持仓Rho      -> rho
         $时间价值    -> time_value   (去掉了前面的 '$')
         到期日       -> expiration_date
       """
    contract: str
    side: str
    total_position: int
    open_price: float
    floating_pnl: float
    floating_pnl_ratio: float
    quoted_pnl: float
    actual_margin: float
    capital_ratio: float
    manual_stop_loss: str
    manual_stop_loss: str
    manual_take_profit: str
    stop_loss_volume: str
    auto_stop_loss: str
    auto_take_profit: str
    position_value: float
    position_type: str
    delta: float
    gamma: float
    theta: float
    vega: float
    rho: float
    time_value: float
    expiration_date: str

@dataclass
class StockRecord:
    index: int                 # 序号
    security_code: str         # 证券代码
    security_name: str         # 证券名称
    balance: float             # 股票余额
    available: float           # 可用余额
    frozen: float              # 冻结数量
    cost_price: float          # 成本价
    market_price: float        # 市价
    profit_loss: float         # 盈亏
    profit_loss_pct: float     # 盈亏比例(%)
    market_value: float        # 市值
    bought_today: float        # 当日买入
    sold_today: float          # 当日卖出
    market: str                # 交易市场

    @classmethod
    def from_dict(cls, d: dict) -> "StockRecord":
        return cls(
            index           = int(d.get('序号', 0)),
            security_code   = d.get('证券代码', ''),
            security_name   = d.get('证券名称', ''),
            balance         = float(d.get('股票余额', 0)),
            available       = float(d.get('可用余额', 0)),
            frozen          = float(d.get('冻结数量', 0)),
            cost_price      = float(d.get('成本价', 0)),
            market_price    = float(d.get('市价', 0)),
            profit_loss     = float(d.get('盈亏', 0)),
            profit_loss_pct = float(d.get('盈亏比例(%)', 0)),
            market_value    = float(d.get('市值', 0)),
            bought_today    = float(d.get('当日买入', 0)),
            sold_today      = float(d.get('当日卖出', 0)),
            market          = d.get('交易市场', '')
        )

@dataclass
class PendingOrder:
    market: str              # 交易市场 (e.g. "深圳Ａ股")
    contract_number: str     # 合同编号 (e.g. "4923263139")
    remark: str              # 备注 (e.g. "未成交")
    order_price: float       # 委托价格 (e.g. 11.020)
    order_quantity: int      # 委托数量 (e.g. 500)
    order_time: str          # 委托时间 (e.g. "16:11:59")
    average_price: float     # 成交均价 (e.g. 0.000)
    trade_quantity: int      # 成交数量 (e.g. 0)
    cancel_quantity: int     # 撤消数量 (e.g. 0)
    operation: str           # 操作 (e.g. "买入")
    symbol: str              # 证券代码 (e.g. "000001")
    security_name: str       # 证券名称 (e.g. "平安银行")