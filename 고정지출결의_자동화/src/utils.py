"""공통 유틸리티"""
import os
import logging
from datetime import datetime
from pathlib import Path


LOG_DIR = Path(__file__).parent.parent / "logs"


def setup_logger(name: str = "expense") -> logging.Logger:
    """로거 설정"""
    LOG_DIR.mkdir(exist_ok=True)

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    # 파일 핸들러
    today = datetime.now().strftime("%Y%m%d")
    fh = logging.FileHandler(LOG_DIR / f"{today}_{name}.log", encoding="utf-8")
    fh.setLevel(logging.INFO)

    # 콘솔 핸들러
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)

    formatter = logging.Formatter("[%(asctime)s] %(levelname)s - %(message)s", "%H:%M:%S")
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


def ensure_dirs():
    """필요한 디렉토리 생성"""
    dirs = [
        Path(__file__).parent.parent / "logs",
        Path(__file__).parent.parent / "screenshots",
        Path(__file__).parent.parent / "attachments",
    ]
    for d in dirs:
        d.mkdir(exist_ok=True)
