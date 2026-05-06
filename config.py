import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
PRINT_TEMPLATES_DIR = os.path.join(BASE_DIR, "print_templates")
USER_DATA_DIR = os.path.join(BASE_DIR, "user_data")

PRESET_MODELS = {
    "cainiao_es": {
        "name": "CAINIAO ES (西班牙菜鸟)",
    },
    "nl": {
        "name": "NL (荷兰仓)",
    },
    "overview": {
        "name": "Inbound/Outbound Overview (综合总览)",
    },
    "byd_car": {
        "name": "BYD乘用车 Inbound&Outbound (乘用车信息汇总)",
    },
}

# Backward-compatible alias
MODELS = PRESET_MODELS


def get_user_dir(uid: str) -> str:
    d = os.path.join(USER_DATA_DIR, uid)
    os.makedirs(d, exist_ok=True)
    return d


def get_user_db_path(uid: str) -> str:
    return os.path.join(get_user_dir(uid), "byd_search.db")


def get_user_upload_dir(uid: str) -> str:
    d = os.path.join(get_user_dir(uid), "uploads")
    os.makedirs(d, exist_ok=True)
    return d
