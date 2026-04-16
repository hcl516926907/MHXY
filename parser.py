import re
from collections import Counter
from typing import List, Optional, Tuple

from ocr_corrections import normalize_ocr_text

# Chinese characters, optionally joined by middle dot "·" (e.g. 神行飞剑·精华)
_CHINESE_ONLY = re.compile(r"^[\u4e00-\u9fff\u3400-\u4dbf\u00b7·]+$")

# UI labels that are NOT item names (quality badges, category tags, etc.)
_UI_LABELS = {"珍品", "精品", "普通", "极品", "套装", "单价"}

# Price: digits optionally separated by commas or periods (OCR sometimes
# misreads "," as "."), e.g. "188,666" / "188.666" / "1888666"
# Must be at least 3 digits total to exclude noise like "1" or "12"
_PRICE_PATTERN = re.compile(r"^[\d,\.]+$")
_MIN_PRICE_DIGITS = 3


def parse_page(
    text_results: List[Tuple[str, float]],
) -> Tuple[Optional[str], List[int]]:
    """
    Parse OCR results from one marketplace page.

    Returns:
        item_name: most frequent Chinese-only text block (= the item being sold)
        prices: list of integer prices extracted from numeric blocks
    """
    item_name_candidates: List[str] = []
    prices: List[int] = []

    for text, _conf in text_results:
        text = normalize_ocr_text(text.strip())
        if not text:
            continue

        # --- Price detection ---
        if _PRICE_PATTERN.match(text):
            digits_only = text.replace(",", "").replace(".", "")
            if len(digits_only) >= _MIN_PRICE_DIGITS:
                try:
                    prices.append(int(digits_only))
                except ValueError:
                    pass
            continue

        # --- Item name detection ---
        if _CHINESE_ONLY.match(text) and len(text) >= 1 and text not in _UI_LABELS:
            item_name_candidates.append(text)

    # The item name is whichever Chinese phrase appears most often on this page
    item_name: Optional[str] = None
    if item_name_candidates:
        counter = Counter(item_name_candidates)
        item_name = counter.most_common(1)[0][0]

    return item_name, prices
