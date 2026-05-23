# coding:utf8
"""
获取股票日K线数据并生成K线图+技术指标图（PNG图片）
技术指标：MACD、KDJ、RSI、BOLL、DMI、EXPMA
使用腾讯日K线接口，支持前复权
"""
import os
import re
import json
import argparse
import datetime
import requests
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.font_manager as fm
from matplotlib.lines import Line2D

# ============ 中文字体设置 ============
def _find_chinese_font():
    """查找系统中可用的中文字体"""
    chinese_fonts = ["SimHei", "Microsoft YaHei", "SimSun", "FangSong", "KaiTi",
                     "WenQuanYi Micro Hei", "Noto Sans CJK SC", "Source Han Sans CN"]
    available = {f.name for f in fm.fontManager.ttflist}
    for font in chinese_fonts:
        if font in available:
            return font
    for f in fm.fontManager.ttflist:
        if any(kw in f.name.lower() for kw in ["hei", "yahei", "song", "fang", "kai", "cjk", "chinese", "noto sans sc"]):
            return f.name
    return None

_CHINESE_FONT = _find_chinese_font()
if _CHINESE_FONT:
    plt.rcParams["font.sans-serif"] = [_CHINESE_FONT] + plt.rcParams["font.sans-serif"]
    plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["axes.unicode_minus"] = False

# ============ 市场前缀映射 ============
MARKET_PREFIX = {
    "6": "sh", "0": "sz", "3": "sz", "688": "sh", "4": "bj", "8": "bj",
}

def _get_market_prefix(code):
    code = code.strip()
    if code[:2] in ("sh", "sz", "bj"):
        return code[:2], code[2:]
    if code.startswith("688"):
        return "sh", code
    prefix = MARKET_PREFIX.get(code[0], "sh")
    return prefix, code

# ============ 数据获取 ============
def fetch_day_kline(stock_code, days=300):
    """从腾讯接口获取日K线数据（前复权）"""
    prefix, num = _get_market_prefix(stock_code)
    full_code = f"{prefix}{num}"
    url = f"http://web.ifzq.gtimg.cn/appstock/app/fqkline/get?_var=kline_dayqfq&param={full_code},day,,,{days},qfq"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    print(f"正在获取 {full_code} 日K线数据（{days}天）...")
    r = requests.get(url, headers=headers, timeout=10)
    match = re.search(r"=(.*)", r.text)
    if not match:
        print(f"解析数据失败: {full_code}")
        return pd.DataFrame(), full_code
    data = json.loads(match.group(1))
    if data.get("code") != 0:
        print(f"接口返回错误: {data.get('msg')}")
        return pd.DataFrame(), full_code
    stock_data = data["data"].get(full_code, {})
    kline_list = stock_data.get("qfqday") or stock_data.get("day", [])
    if not kline_list:
        print(f"未获取到K线数据: {full_code}")
        return pd.DataFrame(), full_code
    # 数据列数可能不一致（6列或7列），统一只取前6列
    kline_list = [row[:6] for row in kline_list]
    df = pd.DataFrame(kline_list, columns=["date", "open", "close", "high", "low", "volume"])
    df["date"] = pd.to_datetime(df["date"])
    for col in ["open", "close", "high", "low"]:
        df[col] = df[col].astype(float)
    df["volume"] = df["volume"].astype(float)
    df = df.set_index("date").sort_index()
    print(f"获取到 {len(df)} 条K线数据")
    return df, full_code

def get_stock_name(stock_code):
    prefix, num = _get_market_prefix(stock_code)
    full_code = f"{prefix}{num}"
    try:
        import easyquotation
        q = easyquotation.use("sina")
        data = q.real(full_code, prefix=True)
        return data.get(full_code, {}).get("name", full_code)
    except Exception:
        return full_code

# ============ 技术指标计算 ============

def calc_ema(series, period):
    return series.ewm(span=period, adjust=False).mean()

def calc_macd(df, fast=12, slow=26, signal=9):
    ema_fast = calc_ema(df["close"], fast)
    ema_slow = calc_ema(df["close"], slow)
    dif = ema_fast - ema_slow
    dea = calc_ema(dif, signal)
    macd_bar = 2 * (dif - dea)
    return dif, dea, macd_bar

def calc_kdj(df, n=9, m1=3, m2=3):
    low_n = df["low"].rolling(window=n).min()
    high_n = df["high"].rolling(window=n).max()
    rsv = (df["close"] - low_n) / (high_n - low_n) * 100
    rsv = rsv.fillna(50)
    k = rsv.ewm(com=m1 - 1, adjust=False).mean()
    d = k.ewm(com=m2 - 1, adjust=False).mean()
    j = 3 * k - 2 * d
    return k, d, j

def calc_rsi(df, periods=None):
    if periods is None:
        periods = [6, 12, 24]
    result = {}
    for n in periods:
        delta = df["close"].diff()
        gain = delta.where(delta > 0, 0.0)
        loss = (-delta).where(delta < 0, 0.0)
        avg_gain = gain.ewm(alpha=1 / n, adjust=False).mean()
        avg_loss = loss.ewm(alpha=1 / n, adjust=False).mean()
        rs = avg_gain / avg_loss
        result[n] = 100 - (100 / (1 + rs))
    return result

def calc_boll(df, n=20, k=2):
    mid = df["close"].rolling(window=n).mean()
    std = df["close"].rolling(window=n).std()
    upper = mid + k * std
    lower = mid - k * std
    return upper, mid, lower

def calc_dmi(df, n=14, m=6):
    high = df["high"]
    low = df["low"]
    close = df["close"]
    tr1 = high - low
    tr2 = abs(high - close.shift(1))
    tr3 = abs(low - close.shift(1))
    tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    up_move = high - high.shift(1)
    down_move = low.shift(1) - low
    plus_dm = np.where((up_move > down_move) & (up_move > 0), up_move, 0.0)
    minus_dm = np.where((down_move > up_move) & (down_move > 0), down_move, 0.0)
    plus_dm = pd.Series(plus_dm, index=df.index)
    minus_dm = pd.Series(minus_dm, index=df.index)
    atr = tr.ewm(alpha=1 / n, adjust=False).mean()
    plus_di = 100 * plus_dm.ewm(alpha=1 / n, adjust=False).mean() / atr
    minus_di = 100 * minus_dm.ewm(alpha=1 / n, adjust=False).mean() / atr
    dx = 100 * abs(plus_di - minus_di) / (plus_di + minus_di)
    adx = dx.ewm(alpha=1 / m, adjust=False).mean()
    adxr = (adx + adx.shift(m)) / 2
    return plus_di, minus_di, adx, adxr

def calc_expma(df, periods=None):
    if periods is None:
        periods = [12, 50]
    result = {}
    for n in periods:
        result[n] = calc_ema(df["close"], n)
    return result

# ============ 绘图 ============

def plot_kline(stock_code, days=300, output_dir=None, ma_list=None,
               show_indicators=None):
    """生成K线图+技术指标图并保存为PNG"""
    if ma_list is None:
        ma_list = [5, 10, 20, 60]
    if show_indicators is None:
        show_indicators = ["macd", "kdj", "rsi", "boll", "dmi", "expma"]
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(__file__))

    df, full_code = fetch_day_kline(stock_code, days)
    if df.empty:
        return None

    stock_name = get_stock_name(stock_code)
    ohlcv = df[["open", "high", "low", "close", "volume"]].copy()

    # ========= 计算指标 =========
    for ma in ma_list:
        ohlcv[f"ma{ma}"] = ohlcv["close"].rolling(window=ma).mean()
    dif, dea, macd_bar = calc_macd(ohlcv)
    k_val, d_val, j_val = calc_kdj(ohlcv)
    rsi_dict = calc_rsi(ohlcv)
    boll_upper, boll_mid, boll_lower = calc_boll(ohlcv)
    plus_di, minus_di, adx, adxr = calc_dmi(ohlcv)
    expma_dict = calc_expma(ohlcv)

    # ========= 最新价和涨跌幅 =========
    last_close = ohlcv["close"].iloc[-1]
    prev_close = ohlcv["close"].iloc[-2] if len(ohlcv) > 1 else last_close
    change_pct = ((last_close / prev_close) - 1) * 100

    # ========= 子图布局 =========
    # 需要独立子图的指标
    sub_indicators = [ind for ind in show_indicators if ind in ("macd", "kdj", "rsi", "dmi")]
    # boll 和 expma 画在主图上
    show_boll = "boll" in show_indicators
    show_expma = "expma" in show_indicators

    n_panels = 2 + len(sub_indicators)  # 主图 + 成交量 + 指标子图
    heights = [5, 1.5] + [1.5] * len(sub_indicators)

    fig, axes = plt.subplots(
        n_panels, 1, figsize=(16, 4 + 2.5 * n_panels),
        gridspec_kw={"height_ratios": heights},
        sharex=True,
    )
    if n_panels == 1:
        axes = [axes]
    fig.subplots_adjust(hspace=0.08, left=0.06, right=0.97, top=0.94, bottom=0.04)

    ax_main = axes[0]
    ax_vol = axes[1]
    ax_indicators = axes[2:]

    dates = ohlcv.index
    x = np.arange(len(dates))

    # ========= 绘制K线 =========
    for i in range(len(ohlcv)):
        o = ohlcv["open"].iloc[i]
        h = ohlcv["high"].iloc[i]
        l = ohlcv["low"].iloc[i]
        c = ohlcv["close"].iloc[i]
        color = "red" if c >= o else "green"
        ax_main.plot([x[i], x[i]], [l, h], color=color, linewidth=0.6)
        body_bottom = min(o, c)
        body_height = abs(c - o) if abs(c - o) > 0.001 else 0.001
        rect = plt.Rectangle((x[i] - 0.3, body_bottom), 0.6, body_height,
                              facecolor=color, edgecolor=color, linewidth=0.5)
        ax_main.add_patch(rect)

    # ========= 均线 =========
    ma_colors = ["#FF9800", "#2196F3", "#9C27B0", "#00BCD4", "#FF5722", "#795548"]
    for i, ma in enumerate(ma_list):
        col = f"ma{ma}"
        if col in ohlcv.columns and len(ohlcv) >= ma:
            ax_main.plot(x, ohlcv[col], color=ma_colors[i % len(ma_colors)],
                         linewidth=0.9, label=f"MA{ma}")

    # ========= BOLL =========
    if show_boll:
        ax_main.plot(x, boll_upper, color="#FF5722", linewidth=0.8, linestyle="--", label="BOLL上轨")
        ax_main.plot(x, boll_mid, color="#2196F3", linewidth=0.8, label="BOLL中轨")
        ax_main.plot(x, boll_lower, color="#FF5722", linewidth=0.8, linestyle="--", label="BOLL下轨")
        ax_main.fill_between(x, boll_lower, boll_upper, alpha=0.06, color="#2196F3")

    # ========= EXPMA =========
    if show_expma:
        expma_colors = ["#E91E63", "#009688"]
        for i, (n, series) in enumerate(expma_dict.items()):
            ax_main.plot(x, series, color=expma_colors[i % len(expma_colors)],
                         linewidth=1.0, linestyle="-.", label=f"EXPMA{n}")

    ax_main.legend(loc="upper left", fontsize=8, ncol=4, framealpha=0.7)
    ax_main.set_ylabel("价格", fontsize=9)
    ax_main.grid(True, alpha=0.3)
    ax_main.set_xlim(x[0] - 1, x[-1] + 1)

    sign = "+" if change_pct >= 0 else ""
    title = f"{stock_name}({full_code})  现价:{last_close:.2f}  涨跌幅:{sign}{change_pct:.2f}%"
    ax_main.set_title(title, fontsize=13, fontweight="bold", pad=10)

    # ========= 成交量 =========
    vol = ohlcv["volume"]
    vol_colors = ["red" if ohlcv["close"].iloc[i] >= ohlcv["open"].iloc[i] else "green"
                  for i in range(len(ohlcv))]
    ax_vol.bar(x, vol, color=vol_colors, width=0.6, alpha=0.8)
    ax_vol.set_ylabel("成交量", fontsize=8)
    ax_vol.grid(True, alpha=0.3)
    ax_vol.tick_params(labelsize=7)

    # ========= 指标子图 =========
    for idx, ind in enumerate(sub_indicators):
        ax = ax_indicators[idx]

        if ind == "macd":
            ax.plot(x, dif, color="#2196F3", linewidth=0.9, label="DIF")
            ax.plot(x, dea, color="#FF9800", linewidth=0.9, label="DEA")
            macd_colors = ["red" if v >= 0 else "green" for v in macd_bar]
            for i in range(len(x)):
                ax.bar(x[i], macd_bar.iloc[i], color=macd_colors[i], width=0.6, alpha=0.7)
            ax.axhline(y=0, color="gray", linewidth=0.5)
            ax.set_ylabel("MACD", fontsize=8)
            ax.legend(loc="upper left", fontsize=7, ncol=3, framealpha=0.7)

        elif ind == "kdj":
            ax.plot(x, k_val, color="#2196F3", linewidth=0.9, label="K")
            ax.plot(x, d_val, color="#FF9800", linewidth=0.9, label="D")
            ax.plot(x, j_val, color="#9C27B0", linewidth=0.9, label="J")
            ax.axhline(y=80, color="gray", linewidth=0.5, linestyle="--", alpha=0.5)
            ax.axhline(y=20, color="gray", linewidth=0.5, linestyle="--", alpha=0.5)
            ax.set_ylabel("KDJ", fontsize=8)
            ax.legend(loc="upper left", fontsize=7, ncol=3, framealpha=0.7)

        elif ind == "rsi":
            rsi_colors = ["#2196F3", "#FF9800", "#9C27B0"]
            for i, (n, series) in enumerate(rsi_dict.items()):
                ax.plot(x, series, color=rsi_colors[i % len(rsi_colors)],
                        linewidth=0.9, label=f"RSI{n}")
            ax.axhline(y=70, color="gray", linewidth=0.5, linestyle="--", alpha=0.5)
            ax.axhline(y=30, color="gray", linewidth=0.5, linestyle="--", alpha=0.5)
            ax.set_ylabel("RSI", fontsize=8)
            ax.legend(loc="upper left", fontsize=7, ncol=3, framealpha=0.7)

        elif ind == "dmi":
            ax.plot(x, plus_di, color="#2196F3", linewidth=0.9, label="+DI")
            ax.plot(x, minus_di, color="#FF9800", linewidth=0.9, label="-DI")
            ax.plot(x, adx, color="#9C27B0", linewidth=0.9, label="ADX")
            ax.plot(x, adxr, color="#00BCD4", linewidth=0.9, linestyle="--", label="ADXR")
            ax.axhline(y=25, color="gray", linewidth=0.5, linestyle="--", alpha=0.5)
            ax.set_ylabel("DMI", fontsize=8)
            ax.legend(loc="upper left", fontsize=7, ncol=4, framealpha=0.7)

        ax.grid(True, alpha=0.3)
        ax.tick_params(labelsize=7)

    # ========= X轴日期 =========
    # 只显示部分日期标签，避免拥挤
    step = max(1, len(dates) // 20)
    tick_positions = x[::step]
    tick_labels = [d.strftime("%m-%d") for d in dates[::step]]
    for ax in axes:
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(tick_labels, fontsize=7, rotation=30)

    # ========= 保存 =========
    now = datetime.datetime.now()
    filename = os.path.join(output_dir, f"{full_code}_K线_{now.strftime('%Y%m%d_%H%M%S')}.png")
    fig.savefig(filename, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"K线图已保存: {filename}")
    return filename


if __name__ == "__main__":
    ALL_INDICATORS = ["macd", "kdj", "rsi", "boll", "dmi", "expma"]
    parser = argparse.ArgumentParser(description="股票日K线图+技术指标生成工具")
    parser.add_argument("stock_code", type=str, help="股票代码，如 600000 或 sh600000")
    parser.add_argument("--days", type=int, default=300, help="K线天数，默认300（建议>=200以保证指标计算充分）")
    parser.add_argument("--ma", type=str, default="5,10,20,60", help="均线周期，逗号分隔，默认 5,10,20,60")
    parser.add_argument("--indicators", type=str, default="macd,kdj,rsi,boll,dmi,expma",
                        help="技术指标，逗号分隔，可选: macd,kdj,rsi,boll,dmi,expma")
    parser.add_argument("--output_dir", type=str, default=None, help="输出目录")
    args = parser.parse_args()

    ma_list = [int(x) for x in args.ma.split(",")]
    indicators = [x.strip() for x in args.indicators.split(",")]
    plot_kline(args.stock_code, days=args.days, output_dir=args.output_dir,
               ma_list=ma_list, show_indicators=indicators)
    print("完成！")
