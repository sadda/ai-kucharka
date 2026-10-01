import logging
import os
from typing import Any

from ..core.data import Data, ExcelFile, InfoStr, WordFile
from .config import ATTRIBUTES


class Company123ExcelFile(ExcelFile):
    def is_paper(self) -> bool:
        return True


class Company123Data(Data):
    excel_file_cls = Company123ExcelFile
    word_file_cls = WordFile

    def __init__(self, root: str, order_name: str, logger: logging.Logger) -> None:
        super().__init__(root, order_name, logger)
        self.info: dict[str, list[InfoStr]] = {attribute: [] for attribute in ATTRIBUTES}

    def __getstate__(self) -> dict[str, Any]:
        state = super().__getstate__()
        state["info"] = self.info
        return state

    def __setstate__(self, state: dict[str, Any]) -> None:
        super().__setstate__(state)
        self.info = state["info"]

    def extract_information(self) -> None:
        for file in self.files:
            if not os.path.basename(file.name).startswith("Zakazka2026"):
                continue
            file.load_tables()
            for attribute in ATTRIBUTES:
                self.info[attribute] = self.info[attribute] + [InfoStr(x, file.name) for x in file.get_info([attribute])]
