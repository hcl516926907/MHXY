"""
梦幻西游跨服套利分析 —— 每日运行脚本

用法：
  python daily_analysis.py
  启动后输入要分析的日期（默认今天），程序读取该日期及之前所有日期的 data/XXXXXX/ CSV，
  将结果写入 analysis/XXXXXX/。

逻辑：
  聚合所有已采集数据：对每件商品，将所有服务器在相同开服天数下的价格合并：
    avg_price = 所有服务器均价的均值
    min_price = 所有服务器最低价的最小值
  以心动服当前开服天数为买入天数，在所有已采集的卖出天数中寻找套利机会：
    条件：卖出天数 - 买入天数 > 冻结天数
    冻结期规则：
      高级兽诀、高级内丹：30 天
      其余商品：           7 天
  对每对（买入天数，卖出天数），计算：
    min 场景：以买入天数的 min_price 买入，卖出天数的 min_price 卖出
    avg 场景：以买入天数的 avg_price 买入，卖出天数的 avg_price 卖出
  卖出收入扣除 9% 交易税。

输出（保存在 analysis/XXXXXX/）：
  price_analysis.csv          —— 所有套利组合
  tradeable_opportunities.csv —— 同上（所有组合均已满足交易条件），按 min 收益率降序排列

模块说明：
  arbitrage.py   —— 套利计算 + 每日 xlsx 报表
  history.py     —— top_opportunities_history 历史汇总
  price_trend.py —— 价格趋势折线图报表
  paths.py       —— 共享路径常量
"""

import os
import glob
import re
from datetime import datetime

import pandas as pd

from paths import DATA_ROOT, ANALYSIS_ROOT
from arbitrage import load_all_source_csvs, aggregate_market, build_analysis, export_xlsx
from history import build_history_xlsx
from price_trend import build_price_trend_xlsx


def prompt_date() -> str:
    """提示用户输入日期，默认为今天。返回 YYYYMMDD 格式字符串。"""
    today = datetime.now().strftime("%Y%m%d")
    while True:
        raw = input(f"请输入分析日期（格式 YYYYMMDD，直接回车默认 {today}）：").strip()
        if raw == "":
            return today
        if re.fullmatch(r"\d{8}", raw):
            return raw
        print("  格式错误，请输入 8 位数字，例如 20260403。")


def main():
    date_str = prompt_date()
    print(f"\n分析日期：{date_str}")

    print("读取原始数据（当前及历史）…")
    df = load_all_source_csvs(date_str)
    servers = sorted(df["server_name"].unique())
    print(f"  共 {len(df)} 行，涉及服务器：{servers}")

    print("聚合市场快照…")
    market = aggregate_market(df)
    days_list = sorted(market["days_open"].unique())
    print(f"  聚合后：{len(market)} 条记录，开服天数范围 {days_list[0]}~{days_list[-1]} 天")

    # 找出心动服在 date_str 当天的开服天数（需在 build_analysis 前确定）
    xindong_days = None
    today_dir = os.path.join(DATA_ROOT, date_str)
    if os.path.isdir(today_dir):
        for path in glob.glob(os.path.join(today_dir, "*.csv")):
            try:
                tmp = pd.read_csv(path, encoding="utf-8-sig")
                tmp_xd = tmp[tmp["server_name"] == "心动"]
                if not tmp_xd.empty:
                    val = pd.to_numeric(tmp_xd["days_open"].iloc[0], errors="coerce")
                    if not pd.isna(val):
                        xindong_days = int(val)
                        break
            except Exception:
                continue
    if xindong_days is not None:
        print(f"  心动服当前开服天数：{xindong_days} 天，仅以此天数作为买入基准计算套利。")
    else:
        print("  [警告] 未找到心动服当前开服天数，将枚举所有买入天数。")

    print("计算套利分析…")
    result = build_analysis(market, fixed_buy_day=xindong_days)

    out_dir = os.path.join(ANALYSIS_ROOT, date_str)
    os.makedirs(out_dir, exist_ok=True)

    # 完整明细（买入天数已限定为心动当前开服天数）
    out_detail = os.path.join(out_dir, "price_analysis.csv")
    result.to_csv(out_detail, index=False, encoding="utf-8-sig")
    print(f"完整明细已保存：analysis/{date_str}/price_analysis.csv（{len(result)} 行）")

    # 套利机会：每商品保留 min 收益率最高的一条
    tradeable = (result
                 .sort_values("收益率%_min", ascending=False)
                 .drop_duplicates(subset="商品名", keep="first")
                 .reset_index(drop=True))
    out_trade = os.path.join(out_dir, "tradeable_opportunities.csv")
    tradeable.to_csv(out_trade, index=False, encoding="utf-8-sig")
    print(f"套利机会已保存：analysis/{date_str}/tradeable_opportunities.csv（{len(tradeable)} 行）")

    # 精简 xlsx 报表
    export_xlsx(tradeable, out_dir, date_str)
    print(f"精简报表已保存：analysis/{date_str}/top_opportunities.xlsx")

    # 控制台摘要
    if not tradeable.empty:
        print("\n── 套利机会（min 收益率前 20）──")
        cols = ["商品名", "分类", "买入天数", "卖出天数", "天数差",
                "买入价(最低)", "卖出价(最低)", "利润_min", "收益率%_min",
                "买入价(均价)", "卖出价(均价)", "利润_avg", "收益率%_avg"]
        print(tradeable[cols].head(20).to_string(index=False))
    else:
        print("\n暂无套利机会。")

    # 更新历史汇总
    print("\n更新历史汇总报表…")
    build_history_xlsx()

    # 价格趋势折线图报表
    print("\n生成价格趋势报表…")
    build_price_trend_xlsx(df)


if __name__ == "__main__":
    main()
