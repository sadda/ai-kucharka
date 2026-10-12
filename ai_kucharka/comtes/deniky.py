from __future__ import annotations

import logging
from collections.abc import Sequence
from typing import Any

from ..core.data import Data, ExcelFile, WordFile
from ..utils import clean_value, collapse_whitespace, remove_whiteshape, split_unit, to_lower

# Key-value blocks of the print log; the labels of the first key anchor the block
header_labels = {
    "number": ["Číslo"],
    "date": ["Datum"],
    "project": ["Projekt"],
    "project_lead": ["Řešitel projektu"],
    "machine": ["Stroj"],
    "wire_diameter": ["Průměr drátu"],
    "head": ["Hlava"],
    "print_time": ["Doba tisku"],
    "printed_by": ["Tisknul"],
    "dimensions": ["Rozměry"],
    "deposition_strategy": ["Depoziční strategie"],
}

information_labels = {
    "information": ["informace"],
}

base_labels = {
    "type": ["Základna"],
    "material": ["Materiál"],
    "dimensions": ["Rozměry"],
    "heat_treatment": ["TZ"],
}

atmosphere_labels = {
    "shielding_gas": ["Ochranný plyn a kvalita"],
    "preheating": ["Předehřev"],
    "preheating_mode": ["Režim předehřevu"],
}

notes_labels = {
    "note": ["Poznámka"],
    "heat_treatment": ["Tepelné zpracování"],
    "metallography": ["Metalografie"],
    "mechanical_tests": ["Mechanické zkoušky"],
}

# Table blocks: header row starts with the anchor, rows are feeders (materials) or parameters (process)
materials_anchor = ["Materiály tisku"]
materials_columns = {
    "designation": ["označení"],
    "manufacturer": ["výrobce"],
    "batch": ["šarže"],
}

process_anchor = ["Proces"]


def normalize_labels(labels: Sequence[Any]) -> list[str]:
    return list(to_lower(collapse_whitespace(remove_whiteshape([str(x) for x in labels]))))


class DenikyExcelFile(ExcelFile):
    def get_materials(self) -> list[dict[str, Any]]:
        blocks = self.find_blocks(materials_anchor)
        if len(blocks) == 0:
            return []
        block = blocks[0]
        header = normalize_labels(list(block.iloc[0, 1:]))
        keys = [next((key for key, values in materials_columns.items() if name in normalize_labels(values)), name) for name in header]
        materials = []
        for row in block.iloc[1:].itertuples(index=False):
            values = {key: clean_value(x) for key, x in zip(keys, row[1:])}
            values = {key: value for key, value in values.items() if value is not None}
            if values:
                materials.append({"feeder": clean_value(row[0]), **values})
        return materials

    def get_process(self) -> list[dict[str, Any]]:
        # One entry per feeder column with parameters given as name, unit and value
        blocks = self.find_blocks(process_anchor)
        if len(blocks) == 0:
            return []
        block = blocks[0]
        process = []
        for j in range(1, block.shape[1]):
            parameters = []
            for i in range(1, block.shape[0]):
                value = clean_value(block.iloc[i, j])
                if value is None:
                    continue
                name, unit = split_unit(str(clean_value(block.iloc[i, 0])))
                parameters.append({"name": name, "unit": unit, "value": value})
            feeder = clean_value(block.iloc[0, j])
            if feeder is not None or parameters:
                process.append({"feeder": feeder, "parameters": parameters})
        return process

    def get_deniky_data(self) -> dict[str, Any] | None:
        header = self.find_labelled_block(header_labels)
        if "number" not in header:
            return None
        header.update(self.find_labelled_block(information_labels))
        return {
            "source": self.name,
            "header": header,
            "materials": self.get_materials(),
            "base": self.find_labelled_block(base_labels),
            "atmosphere": self.find_labelled_block(atmosphere_labels),
            "process": self.get_process(),
            "notes": self.find_labelled_block(notes_labels),
        }


class DenikyData(Data):
    excel_file_cls = DenikyExcelFile
    word_file_cls = WordFile

    def __init__(self, root: str, order_name: str, logger: logging.Logger) -> None:
        super().__init__(root, order_name, logger)
        self.deniky: list[dict[str, Any]] = []

    def use_file(self, file: ExcelFile | WordFile) -> bool:
        return isinstance(file, DenikyExcelFile)

    def extract_file_information(self, file: ExcelFile | WordFile) -> None:
        assert isinstance(file, DenikyExcelFile)
        data = file.get_deniky_data()
        if data is None:
            self.logger.warning("Not a print log", extra={"file": file.name, "type": "not_print_log"})
            return
        data["order"] = self.order_name
        self.deniky.append(data)
