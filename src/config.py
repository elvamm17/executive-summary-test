import os
from pathlib import Path


ROOT_DIR = Path(os.environ.get("WL_BI_ROOT", Path(__file__).resolve().parents[1])).expanduser().resolve()
RAW_DIR = (
    ROOT_DIR / "raw_data"
    if (ROOT_DIR / "raw_data").exists()
    else ROOT_DIR / "raw_data:"
    if (ROOT_DIR / "raw_data:").exists()
    else ROOT_DIR
)
CONFIG_DIR = (
    ROOT_DIR / "config"
    if (ROOT_DIR / "config").exists()
    else ROOT_DIR / "config:"
    if (ROOT_DIR / "config:").exists()
    else ROOT_DIR
)
OUTPUT_DIR = ROOT_DIR / "output"

MASTER_DATA_PATH = OUTPUT_DIR / "master_data.xlsx"
UNMATCHED_SKU_PATH = OUTPUT_DIR / "unmatched_sku.xlsx"

AMAZON_PATTERNS = ("amazon",)
TIKTOK_PATTERNS = ("tiktok", "tik tok")
TEMU_SETTLED_PATTERNS = ("temu", "settled", "已结算", "结算")
TEMU_UNSETTLED_PATTERNS = ("temu", "unsettled", "未结算")
TEMU_ORDER_INFO_PATTERNS = ("订单信息", "order_info", "order info")
TEMU_GUCANG_PATTERNS = ("订单转谷仓", "gucang", "谷仓", "warehouse")
