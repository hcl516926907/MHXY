"""共享路径常量：所有分析模块均从此处导入路径，避免重复定义。"""
import os

ROOT_DIR      = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT     = os.path.join(ROOT_DIR, "data")
ANALYSIS_ROOT = os.path.join(ROOT_DIR, "analysis")
HISTORY_XLSX   = os.path.join(ANALYSIS_ROOT, "top_opportunities_history.xlsx")
DASHBOARD_HTML = os.path.join(ROOT_DIR, "gh-pages", "index.html")


def date_to_dir(root: str, date_str: str) -> str:
    """Return root/YYYY/MM/DD for a YYYYMMDD date string."""
    return os.path.join(root, date_str[:4], date_str[4:6], date_str[6:8])
