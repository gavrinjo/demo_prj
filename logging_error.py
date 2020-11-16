import logging
import sys
import os
from logging.handlers import TimedRotatingFileHandler

FORMATTER = logging.Formatter("\n%(levelname)s - %(asctime)s - %(message)s", datefmt="T(%d.%m.%Y. %H:%M)")
# LOG_FILE = "Logger"


def get_console_handler():
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(FORMATTER)
    return console_handler


def get_file_handler(file):
    file_handler = TimedRotatingFileHandler(f"{file}.log", when="midnight", encoding='utf-8')
    file_handler.setFormatter(FORMATTER)
    return file_handler


def get_logger(logger_name, path=None):
    logger = logging.getLogger(logger_name)

    logger.setLevel(logging.DEBUG)

    logger.addHandler(get_console_handler())
    if path is None:
        logger.addHandler(get_file_handler(logger_name))
    else:
        logger.addHandler(get_file_handler(os.path.join(path, logger_name)))

    logger.propagate = False

    return logger


def get_logger_st():
    logger = logging.getLogger()

    logger.setLevel(logging.DEBUG)

    logger.addHandler(get_console_handler())

    logger.propagate = False

    return logger
