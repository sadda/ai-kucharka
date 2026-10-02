import logging
import os
import warnings
from collections.abc import Sequence

import numpy as np
from pythonjsonlogger import jsonlogger

from .office import convert_office_file


def filter_warnings() -> None:
    warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")


def matching_leading_characters(x: str, y: str) -> int:
    """Returns number of the same leading characters for two strings.

    Args:
        x (str): String 1.
        y (str): String 2.

    Returns:
        Number of the same leading characters.
    """

    i = 0
    max_len = min(len(x), len(y))
    while i < max_len and x[i] == y[i]:
        i += 1
    return i


def most_matching_leading_characters(x: str, ys: Sequence[str], unique: bool = True) -> int | None:
    """Finds the index of string with the most leading characters with another string.

    Args:
        x (str): String 1.
        ys (Sequence[str]): List of strings.
        unique (bool, optional): If true, checks whether the index is unique.

    Returns:
        The index with the most matching leading characters.
    """

    n_matching = [matching_leading_characters(x, y) for y in ys]
    i_max = int(np.argmax(n_matching))
    if unique and np.sum(np.array(n_matching) == n_matching[i_max]) != 1:
        return None
    return i_max


def find_word_files(root, **kwargs) -> tuple[list[str], list[str]]:
    return find_convert_files(root, (".doc", ".docx", ".docm"), ".doc", **kwargs)


def find_excel_files(root, **kwargs) -> tuple[list[str], list[str]]:
    return find_convert_files(root, (".xls", ".xlsx", ".xlsm"), ".xls", **kwargs)


def find_convert_files(root: str, img_extensions: tuple[str, ...], converted_extension: str, **kwargs) -> tuple[list[str], list[str]]:

    original_files = find_files(root, img_extensions=img_extensions, relative_path=False)
    converted_files = list(original_files)
    for i, file_name in enumerate(original_files):
        if file_name.lower().endswith(converted_extension):
            converted_files[i] = convert_office_file(file_name, **kwargs)
    return original_files, converted_files


def find_files(root: str, img_extensions: tuple[str, ...], relative_path: bool = False) -> list[str]:
    """Finds all files in folder and subfolders specified by img_extensions.

    Args:
        root (str): The root folder where to look for images.
        img_extensions (Tuple[str]): Extensions to look for.
        relative_path (bool, optional): Whether relative or absolute paths are returned.

    Returns:
        List of relative paths of the images.
    """

    data = []
    for path, directories, files in os.walk(root):
        for file in files:
            constraint1 = file.lower().endswith(tuple(img_extensions))
            constraint2 = not file.lower().startswith(("~", "."))
            if constraint1 and constraint2:
                if relative_path:
                    file_name = os.path.join(os.path.relpath(path, start=root), file)
                else:
                    file_name = os.path.join(path, file)
                data.append(file_name.replace("/", os.path.sep))
    return data


def unique_str(x: Sequence[str]) -> list[str]:
    arr = np.asarray(x)
    x_lower = np.asarray(x, dtype=str).flatten()
    x_lower = to_lower(remove_whiteshape(x_lower, strip_chars=" \xa0"))
    _, idx = np.unique(x_lower, return_index=True)
    return arr[idx].astype(str).tolist()


def remove_whiteshape(x: np.ndarray | Sequence[str], strip_chars=" :\xa0") -> np.ndarray:
    x_type = np.array(x).dtype
    return np.vectorize(lambda y: y.strip(strip_chars) if isinstance(y, str) else y, otypes=[x_type])(x)


def collapse_whitespace(x: np.ndarray | Sequence[str]) -> np.ndarray:
    x_type = np.array(x).dtype
    return np.vectorize(lambda y: " ".join(y.split()) if isinstance(y, str) else y, otypes=[x_type])(x)


def to_lower(x: np.ndarray | Sequence[str]) -> np.ndarray:
    x_type = np.array(x).dtype
    return np.vectorize(lambda y: y.lower() if isinstance(y, str) else y, otypes=[x_type])(x)


def get_logger(path="log.jsonl", name="mylogger") -> logging.Logger:
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    handler = logging.FileHandler(path, encoding="utf-8")
    formatter = jsonlogger.JsonFormatter()

    handler.setFormatter(formatter)
    logger.addHandler(handler)

    return logger


def close_logger() -> None:
    logging.shutdown()


def close_all_loggers() -> None:
    root = logging.getLogger()

    # Close handlers on root
    for handler in root.handlers[:]:
        handler.close()
        root.removeHandler(handler)

    # Close handlers on all named loggers
    for logger_obj in logging.Logger.manager.loggerDict.values():
        if isinstance(logger_obj, logging.Logger):
            for handler in logger_obj.handlers[:]:
                handler.close()
                logger_obj.removeHandler(handler)
