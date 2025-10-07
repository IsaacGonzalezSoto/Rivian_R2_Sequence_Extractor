"""
Centralized logging configuration for R2_Sequence_Extractor.
"""
import logging
import sys
from pathlib import Path


def setup_logger(name: str, level: int = logging.INFO, log_to_file: bool = False) -> logging.Logger:
    """
    Setup and configure a logger for the application.

    Args:
        name: Logger name (typically __name__ from calling module)
        level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_to_file: Whether to also log to file

    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Avoid duplicate handlers
    if logger.hasHandlers():
        return logger

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)

    # Format: [LEVEL] Module - Message
    formatter = logging.Formatter(
        fmt='[%(levelname)s] %(name)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Optional file handler
    if log_to_file:
        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        file_handler = logging.FileHandler(log_dir / 'extraction.log')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger


def get_logger(name: str, debug: bool = False) -> logging.Logger:
    """
    Get a configured logger instance.

    Args:
        name: Logger name
        debug: Enable debug level logging

    Returns:
        Logger instance
    """
    level = logging.DEBUG if debug else logging.INFO
    return setup_logger(name, level)
