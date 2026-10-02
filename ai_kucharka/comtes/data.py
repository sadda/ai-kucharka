from __future__ import annotations

import logging
import os
import win32com.client
from typing import Any

import numpy as np
import pandas as pd

from ..core.data import Data, ExcelFile, File, InfoComposition, InfoStr, TypeComposition, TypeTest, WordFile
from ..utils import most_matching_leading_characters, remove_whiteshape

keyword_specimen = ["Vzorek", "Vzorky", "Specimen"]
list_of_tests = [
    ("Příloha 12_a – Zkouška tahem – Zkušební těleso kruhové", "tah"),
    ("Příloha 12_b – Zkouška tahem – Zkušební těleso ploché", "tah"),
    ("Příloha 12_c – Zkouška tahem – Zkušební těleso trubka", "tah"),
    ("Příloha 12_d – Zkouška tahem – Zkušební těleso segment z trubky", "tah"),
    ("Příloha 13_a – ČSN EN ISO 148-1 - Zkouška rázem v ohybu", "ohyb"),
    ("Příloha 13_b – ČSN EN ISO 148-1 - Zkouška rázem v ohybu", "ohyb"),
    ("Příloha 13_c – ČSN EN ISO 14556 - Zkouška rázem v ohybu_instrumentovaná zkušební metoda", "ohyb"),
    ("Příloha 13_c – Zkouška rázem v ohybu_instrumentovaná zkušební metoda", "ohyb"),
    ("Příloha 14_b – ČSN EN ISO 5173_Zkouška ohybem - svary", "ohyb"),
    ("Zkouška tahem – Zkušební těleso kruhové", "tah"),
    ("Zkouška tečení – Zkušební těleso kruhové", "teceni"),
    ("Zkouška vysokocyklové únavy", "unava"),
    ("Příloha A_a - Zkouška tlakem", "tlak"),
    ("Příloha B_b – ČSN EN ISO 14556 -  Zkouška rázem v ohybu – ring charpy_instrumentovaná zkušební metoda", "ohyb"),
    ("Příloha 12_b – Zkouška - tříbodový ohyb", "ohyb"),
    ("Results of compress tests", "tlak"),
    ("Results of fatigue tests", "unava"),
    ("Results of four poin bending tests", "ohyb"),
    ("Results of tensile tests", "tah"),
    # TODO: add the following
    ("Zkouška 3PB – Zkušební těleso ploché", "unknown"),
    ("Results of notch toughness tests", "unknown"),
]


class ComtesFile(File):
    """Shared COMTES FHT field-extraction logic, mixed into ComtesWordFile/ComtesExcelFile."""

    def get_material(self) -> list[str]:
        keywords = ["Material", "Materiál"]
        return self.get_info(keywords)

    def get_order(self) -> list[str]:
        keywords = ["Číslo zakázky COMTES FHT", "COMTES FHT job number"]
        return self.get_info(keywords)

    def get_customer(self) -> list[str]:
        keywords = ["Zákazník", "Odběratel", "Customer"]
        return self.get_info(keywords)

    def get_composition(self, min_matches=5) -> list[TypeComposition]:
        compositions = []
        elements = [
            "ag",
            "al",
            "as",
            "b",
            "ba",
            "be",
            "bi",
            "c",
            "ca",
            "cb",
            "cd",
            "ce",
            "cl",
            "co",
            "cr",
            "cu",
            "f",
            "fe",
            "ga",
            "h",
            "hg",
            "in",
            "k",
            "la",
            "li",
            "mg",
            "mn",
            "mo",
            "n",
            "na",
            "ni",
            "nb",
            "o",
            "p",
            "pb",
            "s",
            "sb",
            "se",
            "si",
            "sn",
            "sr",
            "ta",
            "ti",
            "v",
            "w",
            "zn",
            "zr",
        ]
        missing_ok = ["-", "--", "0", "1", "2", "3", "4", "5"]
        for table in self.tables:
            mask = np.isin(table.table_lower, elements)
            i, j = np.where(mask)
            if len(np.unique(table.table_lower[mask])) > min_matches:
                # Extract composition
                valid = i < table.table.shape[0] - 1
                top = table.table[i[valid], j[valid]].astype(object)
                # TODO: make other table orientation as well
                bottom = table.table[i[valid] + 1, j[valid]].astype(object)
                composition = list(zip(top, bottom))
                compositions.append(composition)
                # Print missing elements
                flat = table.table_lower.ravel()
                mask_str = np.vectorize(lambda x: isinstance(x, str))(flat)
                flat = flat[mask_str]
                mask_len = np.isin(np.char.str_len(flat), [1, 2])
                mask_not_in = ~np.isin(flat, elements)
                missing_elements = flat[mask_len & mask_not_in]
                missing_elements = np.setdiff1d(missing_elements, missing_ok)
                if len(missing_elements) > 0:
                    self.logger.warning(
                        "Missing elements",
                        extra={"file": self.name, "type": "missing_elements", "elements": list(missing_elements)},
                    )
        return compositions


class ComtesWordFile(ComtesFile, WordFile):



    def get_cpf_pages(self) -> int:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(os.path.abspath(self.name))
        doc.Repaginate()

        pages = doc.ComputeStatistics(2)

        doc.Close(False)
        word.Quit()

        return pages
    
    def get_cpf_header(self) -> dict[str, Any]:
        labels = {
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

        result = {}
        test_code_heat_number = []
      

        # Projdeme všechny tabulky ve Wordu
        for table in self.tables:
            data = table.table

            for row in data:
                row = [str(value).strip() for value in row]

                for i, cell in enumerate(row):
                    normalized_cell = " ".join(cell.split())

                    if normalized_cell in [
                        "Test code / Heat number",
                        "Označení zkoušky / Číslo tavby",
                    ]:
                        if i + 1 < len(row):
                            value = row[i + 1].strip()


                            if " / E" in value:
                                test_code, heat_number = value.split(" / E", 1)
                                heat_number = "E" + heat_number    

                                test_code_heat_number.append({
                                    "test_code": test_code.strip(),
                                    "heat_number": heat_number.strip(),
                                })

                        continue

                    for key, possible_labels in labels.items():
                        if cell in possible_labels:
                            if i + 1 < len(row):
                                value = row[i + 1].strip()

                                if value:
                                    result[key] = value

        if test_code_heat_number:
            result["test_code_heat_number"] = test_code_heat_number

        result["pages"] = self.get_cpf_pages()    

        return result

    def get_cpf_tests(self) -> dict:
        tensile = []
        impact = []

        for table in self.tables:
            rows = table.table

            # převedení buněk na text
            rows = [
                [str(value).strip() for value in row]
                for row in rows
            ]

            # -----------------
            # TENSILE
            # -----------------
            header = rows[0] if rows else []

            if all(
                column in header
                for column in [
                    "Specimen",
                    "Temp.",
                    "d0",
                    "du",
                    "L0",
                    "Lu",
                    "Rp0,2",
                    "Rm",
                    "Ag",
                    "A",
                    "Z",
                ]
            ):
            
                for row in rows[2:]:
                    if not row[0]:
                        continue

                    tensile.append({
                        "specimen": row[0],
                        "temp": row[1],
                        "d0": row[2],
                        "du": row[3],
                        "L0": row[4],
                        "Lu": row[5],
                        "Rp0,2": row[6],
                        "Rm": row[7],
                        "Ag": row[8],
                        "A": row[9],
                        "Z": row[10],
                    })

            # -----------------
            # IMPACT / CHARPY
            # -----------------
            if len(rows) >= 2:
                header = rows[1]

                if all(
                    column in header
                    for column in [
                        "Specimen",
                        "B",
                        "W",
                        "Ligament",
                        "Temp.",
                        "KU2" if "KU2" in header else "KV2",
                        "Fracture",
                    ]
                ):
                    













                    
                    for row in rows[3:]:
                        if not row[0]:
                            continue

                        impact.append({
                            "specimen": row[0],
                            "B": row[1],
                            "W": row[2],
                            "ligament": row[3],
                            "temp": row[4],
                            ("KU2" if "KU2" in header else "KV2"): row[5],
                            "fracture": row[6],
                        })

        return {
            "tensile": tensile,
            "impact": impact,
    
        }
    
    def get_cpf_methods(self) -> list[dict[str, str]]:
        methods = []
        current_method = None

        for paragraph in self.file.paragraphs:
            text = paragraph.text.strip()

            if not text:
                continue

            if text.startswith("Test method name:"):
                current_method = {
                    "name": text.split(":", 1)[1].strip(),
                }
                methods.append(current_method)

            elif text.startswith("Test method identification:"):
                if current_method is not None:
                    current_method["identification"] = text.split(":", 1)[1].strip()

            elif text.startswith("Test specimen type:"):
                if current_method is not None:
                    current_method["specimen_type"] = text.split(":", 1)[1].strip()

        return methods
        
    def get_cpf_data(self) -> dict:
        return {
            "header": self.get_cpf_header(),
            "methods": self.get_cpf_methods(),
            "tests": self.get_cpf_tests(),
        }
    
class ComtesExcelFile(ComtesFile, ExcelFile):
    def is_paper(self) -> bool:
        base_name = os.path.basename(self.name)
        return "paper" in base_name or "Paper" in base_name

    def get_order(self) -> list[str]:
        order = super().get_order()
        for table in self.tables:
            idx = [isinstance(x, str) and x.startswith("zak") for x in table.table_lower.flatten()]
            order = order + table.table.flatten()[idx].tolist()
        return order


class ComtesData(Data):
    excel_file_cls = ComtesExcelFile
    word_file_cls = ComtesWordFile

    def __init__(
        self,
        root: str,
        order_name: str,
        logger: logging.Logger,
    ) -> None:
        super().__init__(root, order_name, logger)
        self.tests: list[TypeTest] = []
        self.material: list[InfoStr] = []
        self.order: list[InfoStr] = []
        self.customer: list[InfoStr] = []
        self.composition: list[InfoComposition] = []
        self.values_tests: pd.DataFrame = pd.DataFrame()

    def extract_file_information(self, file: ExcelFile | WordFile) -> None:
        assert isinstance(file, ComtesExcelFile | ComtesWordFile)
        self.material = self.material + [InfoStr(x, file.name) for x in file.get_material()]
        self.order = self.order + [InfoStr(x, file.name) for x in file.get_order()]
        self.customer = self.customer + [InfoStr(x, file.name) for x in file.get_customer()]
        self.composition = self.composition + [InfoComposition(x, file.name) for x in file.get_composition()]

    def use_file(self, file: ExcelFile | WordFile) -> bool:
        return not isinstance(file, ComtesExcelFile) or file.is_paper()

    def postprocess_information(self) -> None:
        self.tests = [self._get_info_from_paper(file) for file in self.files]
        self._extract_test_information()

    def _get_info_from_paper(self, file: File) -> TypeTest:
        if not isinstance(file, ComtesExcelFile) or not file.is_paper() or len(file.sheet_names) == 0:
            return [], "", ""

        sheet = file.load_sheet(file.sheet_names[0])
        if len(sheet) == 0:
            return [], "", ""
        col = sheet.columns[0]

        match = next(
            ((name, t) for name, t in list_of_tests if name in sheet.columns),
            None,
        )
        if match is None:
            # TODO: finish
            self.logger.warning("Wrong headers", extra={"file": file.name, "type": "wrong_headers", "columns": list(sheet.columns)})
            return [], "", ""
        test_name, test_type = match

        # Check that the sheet has correct format
        idx = np.where(sheet[col].isin(keyword_specimen))[0]
        if len(idx) != 1:
            self.logger.warning("Specimen not found", extra={"file": file.name, "type": "no_specimen"})
            return [], "", ""

        # TODO: check better in other columns as well
        names = sheet[col].iloc[idx[0] + 1 :]
        names = names[~names.isnull()]
        names = names.astype(str).str.strip().to_numpy().tolist()
        # Ignore all information after keywords
        keywords = ["Zkušební stroj", "L0 ext", "TR", "TF", "Součinitel asymetrie cyklu"]
        keywords = remove_whiteshape(keywords)
        for keyword in keywords:
            names_lower = remove_whiteshape(names)
            idx = np.where(names_lower == keyword)[0]
            if len(idx) == 1:
                names = names[: idx[0]]
        return names, test_name, test_type

    def _extract_test_information(self) -> None:
        result_sheet_name = "Results sheet"
        for file, (names, test_name, test_type) in zip(self.files, self.tests):
            if test_type not in ("tah", "tlak"):
                continue
            for name in names:
                assert isinstance(file, ExcelFile)
                file_name = file.name
                i = self._find_report_index(file_name, name)
                if i is None:
                    continue

                file_results = self.files[i]
                assert isinstance(file_results, ExcelFile)
                if result_sheet_name not in file_results.sheet_names:
                    self.logger.warning("Result sheet not found", extra={"file": file_results.name, "type": "result_sheet_not_found"})
                    continue

                table = file_results.load_sheet(result_sheet_name)
                specimen = table.iloc[1, 1]
                if table.shape[0] < 4 or table.shape[1] < 2 or specimen != "Specimen":
                    self.logger.warning("Results different format", extra={"file": file_results.name, "type": "results_different_format"})
                    continue
                headers = table.iloc[1, 1:].tolist()
                values = table.iloc[3, 1:].tolist()
                values_test = pd.DataFrame([values], columns=headers)
                values_test = values_test.loc[:, ~values_test.columns.duplicated()]
                specimen = str(values_test.loc[0, "Specimen"])
                values_test["Specimen"] = InfoStr(specimen, file_results.name)
                self.values_tests = pd.concat((self.values_tests, values_test))
        self.values_tests = self.values_tests.reset_index(drop=True)

    def _find_report_index(self, file_name, name) -> None | int:
        sep = os.path.sep
        idx = np.where(self.metadata["base_name"] == name)[0]
        if len(idx) == 0:
            # Name was not found
            self.logger.warning("Name not found", extra={"file": file_name, "type": "name_not_found", "test_name": name})
            return None
        elif len(idx) == 1:
            # There is exactly one name
            return idx[0]
        else:
            # There are multiple names
            file_name_split = file_name.split(sep)
            # Check for priority structure
            if file_name_split[-2] == "data":
                file_name_split[-2] = "results"
                file_name_split[-1] = name
                for ext in [".xls", ".xlsx", ".xlsm"]:
                    file_name_new = sep.join(file_name_split) + ext
                    if os.path.exists(file_name_new):
                        i = int(np.where(self.metadata["original_name"].iloc[idx] == file_name_new)[0][0])
                        assert isinstance(i, int)
                        return idx[i]
            # Check for the most matching file name
            i = most_matching_leading_characters(file_name, self.metadata["original_name"].iloc[idx].tolist())
            if i is None:
                self.logger.warning("Multiple names", extra={"file": file_name, "type": "multiple_names", "test_name": name})
                return i
            return idx[i]
