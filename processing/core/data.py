from __future__ import annotations

import logging
import numbers
import os
import pickle
import re
import unicodedata
from abc import ABC, abstractmethod
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any, ClassVar, TypeVar, cast

import numpy as np
import pandas as pd
from docx import Document
from numpy.typing import ArrayLike

from ..utils import (
    find_excel_files,
    find_word_files,
    remove_whiteshape,
    to_lower,
)

_default_logger = logging.getLogger(__name__)
_default_logger.addHandler(logging.NullHandler())


TypeComposition = list[tuple[str, str]]
TypeTest = tuple[list[Any], str, str]

TData = TypeVar("TData", bound="Data")


@dataclass
class Info:
    value: Any
    source: str

    def fix_path(self) -> None:
        self.source = self.source.replace(os.path.sep, "/")


@dataclass
class InfoStr(Info):
    value: str
    source: str

    def normalize_value(self) -> None:
        x = self.value
        if not isinstance(x, str):
            return x

        # Normalize unicode (handles weird composed characters)
        x = unicodedata.normalize("NFKC", x)

        # Replace non-breaking spaces and similar with normal space
        x = x.replace("\xa0", " ")

        # Remove zero-width and control characters
        x = re.sub(r"[\u200B-\u200D\uFEFF]", "", x)

        # Collapse all whitespace to single spaces
        x = re.sub(r"\s+", " ", x)

        self.value = x.strip()

    def to_str(self) -> None:
        if isinstance(self.value, numbers.Number):
            self.value = str(self.value)
        # TODO: should not be here
        import datetime

        if isinstance(self.value, datetime.datetime):
            self.value = "???"


@dataclass
class InfoComposition(Info):
    value: TypeComposition
    source: str


class Table:
    def __init__(self, table: ArrayLike) -> None:

        table = np.asarray(table)
        if table.ndim == 1:
            table = table.reshape(1, -1)

        self.table = remove_whiteshape(table)
        self.table_lower = to_lower(self.table)

    def get_info(self, keywords: Sequence[str]) -> list[str]:
        idx = np.isin(self.table_lower, keywords)
        ii, jj = np.where(idx)
        info = []
        for i, j in zip(ii, jj):
            # TODO: add merged cells
            row = self.table[i, j + 1 :]
            row = row[~pd.isnull(row)]
            if len(row) > 0:
                info.append(row[0])
        return info


class File(ABC):
    def __init__(self, file_name: str, logger: logging.Logger, display_name: str | None = None) -> None:

        if display_name is None:
            display_name = file_name
        self.logger = logger
        self.tables = []
        self._file_name = file_name
        self._display_name = display_name

    @property
    def name(self) -> str:
        return self._display_name

    def get_info(self, keywords: Sequence[str]) -> list[str]:
        keywords_lower = to_lower(remove_whiteshape(keywords))
        info = []
        for table in self.tables:
            info = info + table.get_info(keywords_lower)
        return info


class WordFile(File):
    def __init__(self, file_name: str, logger: logging.Logger, **kwargs) -> None:
        super().__init__(file_name, logger, **kwargs)
        try:
            self.file = Document(self._file_name)
        except Exception:
            self.logger.warning("Unreadable file", extra={"file": self.name, "type": "unreadable_file"})
            self.file = None

    def load_tables(self) -> None:
        if self.file:
            for tab in self.file.tables:
                table = []
                for row in tab.rows:
                    # TODO: possibly skip the whole table
                    try:
                        row = [cell.text for cell in row.cells]
                    except ValueError:
                        continue
                    if any(r != "" for r in row):
                        table.append(row)
                if len(table) > 0:
                    lens = [len(row) for row in table]
                    if len(np.unique(lens)) != 1:
                        max_len = max(lens)
                        table = [list(x) + [np.nan] * (max_len - len(x)) for x in table]
                    self.tables.append(Table(table))

    def close(self) -> None:
        pass


class ExcelFile(File):
    def __init__(self, file_name: str, logger: logging.Logger, **kwargs) -> None:
        """Class for a single Excel file.

        It creates the Excel instance, extracts sheets.
        Does not load the content of the excel file yet.

        Args:
            file_name (str): Name of the Excel file.
        """

        super().__init__(file_name, logger, **kwargs)

        try:
            self.file = pd.ExcelFile(self._file_name)
            self.sheet_names = cast(list[str], self.file.sheet_names)
            assert all(isinstance(name, str) for name in self.sheet_names)
            if len(self.sheet_names) == 0:
                raise Exception("Excel file has no sheets")
        except Exception as e:
            self.logger.warning("Unreadable file", extra={"file": self.name, "type": "unreadable_file", "exception": str(e)})
            self.file = None
            self.sheet_names = []

    def load_tables(self) -> None:
        if self.file and self.sheet_names:
            try:
                sheet = self.load_sheet(self.sheet_names[0])
                self.tables = [Table(sheet)]
            except Exception as e:
                self.logger.warning("Unreadable file", extra={"file": self.name, "type": "unreadable_file", "exception": str(e)})

    def load_sheet(self, sheet_name: str) -> pd.DataFrame:
        """Loads content of a single sheet from the file.

        Args:
            sheet_name (str): Name of the sheet.

        Returns:
            Loaded data.
        """

        return pd.read_excel(self.file, sheet_name, parse_dates=False)

    def load_sheets(self) -> dict[str, pd.DataFrame]:
        """Loads all sheets from the file and put them into a dictionary.

        Returns:
            Loaded data.
        """

        return {sheet_name: self.load_sheet(sheet_name) for sheet_name in self.sheet_names}

    def close(self) -> None:
        if self.file is not None:
            try:
                self.file.close()
            except Exception as e:
                self.logger.warning("File did not close", extra={"file": self.name, "type": "closing_error", "exception": str(e)})


class Data(ABC):
    excel_file_cls: ClassVar[type[ExcelFile]]
    word_file_cls: ClassVar[type[WordFile]]
    # Attributes which are not pickled (open files and logger handlers)
    _runtime_attributes: ClassVar[tuple[str, ...]] = ("files", "logger")

    def __init__(
        self,
        root: str,
        order_name: str,
        logger: logging.Logger,
    ) -> None:

        self.root = root
        self.order_name = order_name
        self.logger = logger
        self.files: list[ExcelFile | WordFile] = []
        self.file_names: list[str] = []
        self.metadata = pd.DataFrame()

    def __getstate__(self) -> dict[str, Any]:
        return {k: v for k, v in self.__dict__.items() if k not in self._runtime_attributes}

    def __len__(self) -> int:
        return len(self.metadata)

    def __setstate__(self, state: dict[str, Any]) -> None:
        self.__dict__.update(state)
        # Initialize runtime-only attributes
        self.files = []
        self.logger = _default_logger

    def extract_files(self, **kwargs) -> None:
        original_names, converted_names = find_excel_files(self.root, **kwargs)
        metadata1 = pd.DataFrame(
            {
                "original_name": original_names,
                "converted_name": converted_names,
                "base_name": [os.path.splitext(os.path.basename(file_name))[0] for file_name in original_names],
                "type": "excel",
            }
        )
        files1 = [self.excel_file_cls(x, self.logger, display_name=y) for x, y in zip(converted_names, original_names)]

        original_names, converted_names = find_word_files(self.root, **kwargs)
        metadata2 = pd.DataFrame(
            {
                "original_name": original_names,
                "converted_name": converted_names,
                "base_name": [os.path.splitext(os.path.basename(file_name))[0] for file_name in original_names],
                "type": "word",
            }
        )
        files2 = [self.word_file_cls(x, self.logger, display_name=y) for x, y in zip(converted_names, original_names)]

        self.files = files1 + files2
        self.file_names = [file.name for file in self.files]
        self.metadata = pd.concat((metadata1, metadata2)).reset_index(drop=True)

    def use_file(self, file: ExcelFile | WordFile) -> bool:
        return True

    def extract_information(self) -> None:
        for file in self.files:
            if not self.use_file(file):
                continue
            file.load_tables()
            self.extract_file_information(file)
        self.postprocess_information()

    @abstractmethod
    def extract_file_information(self, file: ExcelFile | WordFile) -> None:
        """Extract company-specific information from one file."""

    def postprocess_information(self) -> None:
        pass

    @classmethod
    def load(cls: type[TData], path: str, logger: logging.Logger | None = None) -> TData:
        with open(path, "rb") as f:
            obj = pickle.load(f)
        if not isinstance(obj, cls):
            raise TypeError("Loaded object is not of type Data")
        if logger is not None:
            obj.logger = logger
        return obj

    def save(self, path: str) -> None:
        with open(path, "wb") as f:
            pickle.dump(self, f, protocol=pickle.HIGHEST_PROTOCOL)

    def close_files(self) -> None:
        for file in self.files:
            file.close()
