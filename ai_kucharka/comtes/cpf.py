from __future__ import annotations

import logging
import os
from typing import Any

from ..core.data import ExcelFile, WordFile
from .data import ComtesData, ComtesWordFile

header_labels = {
    "title": ["Title", "Název"],
    "order_number": ["Order number", "Číslo objednávky"],
    "test_code": ["Test code", "Označení zkoušky"],
    "job_number": ["Job number", "Číslo zakázky"],
    "material": ["Material", "Materiál"],
    "heat_number": ["Heat number", "Číslo tavby"],
    "drawing_number": ["Drawing number", "Číslo výkresu"],
    "report_number": ["Report number", "Číslo protokolu"],
    "pages": ["Pages", "Počet stran"],
    "appendices": ["Appendices", "Počet příloh"],
    "order_received": ["Order received", "Datum příjmu objednávky"],
    "test_date": ["Test date", "Datum zkoušky"],
    "report_date": ["Report date", "Datum vydání protokolu"],
    "comtes_job_number": [
        "COMTES FHT job number",
        "Číslo zakázky COMTES FHT",
    ],
}

test_code_heat_number_labels = ["Test code / Heat number", "Označení zkoušky / Číslo tavby"]

# Paragraphs "label: value" describing test methods; a new method starts with its name
method_labels = {
    "name": ["Test method name", "Přesný název zkušební metody", "Genaue Benennung der Methode"],
    "identification": ["Test method identification", "Identifikace zkušební metody", "Identifikation der Methode"],
    "specimen_type": ["Test specimen type", "Typ zkušebního tělesa", "Art des Probekörpers"],
}

# Test tables: index of the header row and required columns (a tuple lists alternative names)
test_tables = {
    "tensile": (0, ["Specimen", "Temp.", "d0", "du", "L0", "Lu", "Rp0,2", "Rm", "Ag", "A", "Z"]),
    "impact": (1, ["Specimen", "B", "W", "Ligament", "Temp.", ("KU2", "KV2"), "Fracture"]),
}


class CPFWordFile(ComtesWordFile):
    def get_cpf_header(self) -> dict[str, Any]:
        result = {key: self.find_values_next_to_labels(labels) for key, labels in header_labels.items()}
        result = {key: values for key, values in result.items() if values}

        test_code_heat_number = []
        for value in self.find_values_next_to_labels(test_code_heat_number_labels):
            if " / E" in value:
                test_code, heat_number = value.split(" / E", 1)
                test_code_heat_number.append({"test_code": test_code.strip(), "heat_number": ("E" + heat_number).strip()})
        if test_code_heat_number:
            result["test_code_heat_number"] = test_code_heat_number

        return result

    def get_cpf_tests(self) -> dict[str, list[dict[str, str]]]:
        tests = {}
        for name, (header_row, columns) in test_tables.items():
            tests[name] = [row for df in self.find_rows(columns, header_row=header_row, skip_rows=1) for row in df.to_dict("records")]
        return tests

    def get_cpf_methods(self) -> list[dict[str, str]]:
        methods = []
        for key, value in self.find_labelled_paragraphs(method_labels):
            if key == "name":
                methods.append({"name": value})
            elif methods:
                methods[-1][key] = value
        return methods

    def get_cpf_data(self) -> dict:
        return {
            "source": self.name,
            "header": self.get_cpf_header(),
            "methods": self.get_cpf_methods(),
            "tests": self.get_cpf_tests(),
        }


class CPFData(ComtesData):
    word_file_cls = CPFWordFile

    def __init__(self, root: str, order_name: str, logger: logging.Logger) -> None:
        super().__init__(root, order_name, logger)
        self.cpf: list[dict] = []

    def extract_file_information(self, file: ExcelFile | WordFile) -> None:
        super().extract_file_information(file)
        if isinstance(file, CPFWordFile) and "_CPF_" in os.path.basename(file.name):
            self.cpf.append(file.get_cpf_data())
