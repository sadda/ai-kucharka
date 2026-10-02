import logging
import os

from ..core.data import Data, ExcelFile, InfoStr, WordFile
from .config import ATTRIBUTES


class Company123Data(Data):
    excel_file_cls = ExcelFile
    word_file_cls = WordFile

    def __init__(self, root: str, order_name: str, logger: logging.Logger) -> None:
        super().__init__(root, order_name, logger)
        self.info: dict[str, list[InfoStr]] = {attribute: [] for attribute in ATTRIBUTES}

    def use_file(self, file: ExcelFile | WordFile) -> bool:
        return os.path.basename(file.name).startswith("Zakazka2026")

    def extract_file_information(self, file: ExcelFile | WordFile) -> None:
        for attribute in ATTRIBUTES:
            self.info[attribute] = self.info[attribute] + [InfoStr(x, file.name) for x in file.find_values_next_to_labels([attribute])]
