from __future__ import annotations

import logging
import os
import platform
from typing import Any

from ..core.data import ExcelFile, WordFile
from .data import ComtesData, ComtesWordFile


class CPFWordFile(ComtesWordFile):



    def get_cpf_pages(self) -> int | None:
        if platform.system() != "Windows":
            return None
        import win32com.client

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
    

class CPFData(ComtesData):
    word_file_cls = CPFWordFile

    def __init__(self, root: str, order_name: str, logger: logging.Logger) -> None:
        super().__init__(root, order_name, logger)
        self.cpf: list[dict] = []

    def extract_file_information(self, file: ExcelFile | WordFile) -> None:
        super().extract_file_information(file)
        if isinstance(file, CPFWordFile) and "_CPF_" in os.path.basename(file.name):
            self.cpf.append({"source": file.name, **file.get_cpf_data()})
