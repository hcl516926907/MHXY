import csv
import os
from datetime import datetime
from typing import List, Optional


def _csv_path(output_dir: str, server_name: str) -> str:
    today = datetime.now().strftime("%Y%m%d")
    day_dir = os.path.join(output_dir, today)
    os.makedirs(day_dir, exist_ok=True)
    return os.path.join(day_dir, f"{today}_{server_name}.csv")


_HEADER = ["timestamp", "server_name", "open_date", "item_name", "category",
           "avg_price", "min_price", "count", "prices_raw", "days_open"]


def save_result(
    item_name: Optional[str],
    prices: List[int],
    output_dir: str,
    server_name: str = "未知",
    server_open_dt=None,
    category: str = "",
) -> str:
    """
    Compute average/min price and append one row to today's CSV file.

    CSV columns: timestamp | server_name | open_date | item_name |
                 avg_price | min_price | count | prices_raw | days_open
    """
    os.makedirs(output_dir, exist_ok=True)

    filepath = _csv_path(output_dir, server_name)
    file_exists = os.path.exists(filepath)

    now = datetime.now()
    avg_price = round(sum(prices) / len(prices)) if prices else 0
    min_price = min(prices) if prices else 0
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    prices_raw = "|".join(str(p) for p in sorted(prices))
    open_date = server_open_dt.strftime("%Y-%m-%d") if server_open_dt is not None else ""
    days_open = (now.date() - server_open_dt.date()).days + 1 if server_open_dt is not None else ""

    with open(filepath, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(_HEADER)
        writer.writerow(
            [timestamp, server_name, open_date, item_name or "未知", category,
             avg_price, min_price, len(prices), prices_raw, days_open]
        )

    return filepath


def save_na_result(
    item_name: str,
    output_dir: str,
    server_name: str = "未知",
    server_open_dt=None,
    category: str = "",
) -> str:
    """货架为空时写入一行 NA 记录，与 save_result 共用同一 CSV 文件。"""
    os.makedirs(output_dir, exist_ok=True)

    filepath = _csv_path(output_dir, server_name)
    file_exists = os.path.exists(filepath)

    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    open_date = server_open_dt.strftime("%Y-%m-%d") if server_open_dt is not None else ""
    days_open = (now.date() - server_open_dt.date()).days + 1 if server_open_dt is not None else ""

    with open(filepath, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(_HEADER)
        writer.writerow([timestamp, server_name, open_date, item_name, category,
                         "NA", "NA", 0, "NA", days_open])

    return filepath
