"""Logging setup — rotating file + stdout."""
from __future__ import annotations
import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logger(log_file: Path) -> logging.Logger:
    log_file.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("ice_quote_agent")
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger  # already configured

    fmt = logging.Formatter(
        "%(asctime)s  %(levelname)-7s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = RotatingFileHandler(log_file, maxBytes=5_000_000, backupCount=5)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    return logger
