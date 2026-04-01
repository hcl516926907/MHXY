import csv
import os
from datetime import datetime
from typing import List, Optional


def save_result(
    item_name: Optional[str],
    prices: List[int],
    output_dir: str,
    server_name: str = "未知",
    server_open_dt=None,
) -> str:
    """
    Compute average price and append one row to today's CSV file.

    CSV columns: timestamp | server_name | open_date | item_name | avg_price | count | prices_raw | days_open

    Returns the path of the CSV file written to.
    """
    os.makedirs(output_dir, exist_ok=True)

    today = datetime.now().strftime("%Y%m%d")
    filepath = os.path.join(output_dir, f"prices_{today}.csv")
    file_exists = os.path.exists(filepath)

    now = datetime.now()
    avg_price = round(sum(prices) / len(prices)) if prices else 0
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    prices_raw = "|".join(str(p) for p in sorted(prices))
    open_date = server_open_dt.strftime("%Y-%m-%d") if server_open_dt is not None else ""
    days_open = (now - server_open_dt).days + 1 if server_open_dt is not None else ""

    with open(filepath, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["timestamp", "server_name", "open_date", "item_name", "avg_price", "count", "prices_raw", "days_open"])
        writer.writerow(
            [timestamp, server_name, open_date, item_name or "未知", avg_price, len(prices), prices_raw, days_open]
        )

    return filepath
