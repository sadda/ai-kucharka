from __future__ import annotations

from email import header
import logging
import os
from typing import Any
from unittest import result

import pandas as pd

from ..core.data import Data, ExcelFile, WordFile, to_python_str

header_labels = {
    "title": ["Title", "Název", "Bezeichnung"],
    "order_number": ["Order number", "Číslo objednávky", "Bestellnummer"],
    "test_code": ["Test code", "Označení zkoušky", "Identifikationsnummer der Prüfung"],
    "job_number": ["Job number", "Číslo zakázky", "Auftragsnummer"],
    "material": ["Material", "Materiál", "Werkstoff"],
    "heat_number": ["Heat number", "Číslo tavby", "Schmelzenummer"],
    "drawing_number": ["Drawing number", "Číslo výkresu", "Zeichnungsnummer"],
    "report_number": ["Report number", "Číslo protokolu", "Protokollnummer"],
    "pages": ["Pages", "Počet stran", "Seiten insgesamt"],
    "appendices": ["Appendices", "Počet příloh", "Anlage"],
    "order_received": ["Order received", "Datum příjmu objednávky", "Empfang der Proben am"],
    "test_date": ["Test date", "Datum zkoušky", "Durchführung der Versuche am"],
    "report_date": ["Report date", "Datum vydání protokolu", "Prüfung abgeschlossen am"],
    "comtes_job_number": [
        "COMTES FHT job number",
        "Číslo zakázky COMTES FHT",
        "Auftragsnummer COMTES FHT"
    ],
    "note": ["Note", "Poznámka", "Note"],
}


# Varianta 3:
# Test code / Heat number jsou v jedné buňce.
test_code_heat_number_labels = [
    "Test code / Heat number",
    "Označení zkoušky / Číslo tavby",
    "Test code / Heat No.",
    "Test code; Heat number",
    "Identifikationsnummer der Prüfung / Schmelzennummer",
    "Identifikationsnummer der Prüfung /  Schmelzenummer",
    "Identifikationsnummer der Prütung /Schmelzenummer",
]

test_code_labels = [
    "Identifikationsnummer der Prüfung",
]

heat_number_labels = [
    "Schmelzennummer",
    "Schmelzenummer",
]

# Varianta 2:
# Material / Heat No. jsou v jedné buňce.
material_heat_number_labels = [
    "Material / Heat No.",
    "Material / Heat number",
    "Materiál / Číslo tavby",
    "Werkstoff / Schmelzenummer",
]
job_number_note_labels = [
    "Job number / Note",
    "Auftragsnummer / Note",
]

# NOVÉ – Varianta 5 (Pouze DE)
job_number_drawing_labels = [
    "Auftragsnummer / Zeichnungsnummer",
    "AuftragsNo./ZeichnungsNo.",
]

# NOVÉ – Varianta 5 (Pouze DE)
heat_number_material_labels = [
    "Schmelzennummer / Werkstoff",
    "SchmelzeNo./Werkstoff",
]

# NOVÉ – Varianta 3 (Pouze CZ)
job_number_material_labels = [
    "Číslo zakázky / Materiál",
]

material_drawing_labels = [
    "Material / Drawing No.",
]

# Textový požadavek zákazníka.
# Celý obsah ponecháváme jako jeden string.
customer_requirement_labels = [
    "Data provided by the customer",
    "Data provided by the customer:",
    "Data poskytnutá zákazníkem",
    "Vom Kunden zur Verfügung gestellte Daten",
    
]

written_by_labels = [
    "Written by",
    "Vypracoval",
    "Ausgearbeitet von",
]

approved_by_labels = [
    "Approved by",
    "Schválil",
    "Freigegeben von",
]
method_labels = {
    "name": [
        "Test method name",
        "Přesný název zkušební metody",
        "Genaue Benennung der Methode",
    ],
    "identification": [
        "Test method identification",
        "Identifikace zkušební metody",
        "Identifikation der Methode",
    ],
    "specimen_type": [
        "Test specimen type",
        "Typ zkušebního tělesa",
        "Art des Probekörpers",
    ],
}

test_series_labels = {
    "impact": [
        "Impact bend test",
        "Zkouška rázem v ohybu",
        "Schlagbiegeversuch",
    ],
    "tensile": [
        "Tensile test",
        "Zkouška tahem",
        "Zugprüfung",
    ],
    "hardness": [
        "Brinell hardness test",
        "Zkouška tvrdosti podle Brinella",
    ],
}

test_tables = {
    "tensile": (
        0,
        [
            ("Specimen", "Vzorek", "Probenummer"),
            ("Temp.", "Tepl.", "Temp."),
            "d0",
            "du",
            "L0",
            "Lu",
            "Rp0,2",
            "Rm",
            "Ag",
            "A",
            "Z",
        ],
    ),
    "impact": (
        1,
        [
            ("Specimen", "Vzorek", "Probenummer", "Probennummer"),
            ("B",),
            ("W",),
            ("Ligament",),
            ("Temp.", "Tepl.", "Teplota"),
            ("KU2", "KV2"),
            ("Fracture", "Rozlomení", "Fraktur"),
        ],
    ),
    "hardness": (
        0,
        [
            ("Specimen", "Vzorek"),
            ("Hardness measurements", "Měření tvrdosti"),
            ("Average", "Průměr"),
        ],
    ),
}


class CPFWordFile(WordFile):
    def get_cpf_header(self) -> dict[str, Any]:
        result = {
            key: self.find_values_next_to_labels(labels)
            for key, labels in header_labels.items()
        }

        result = {
            key: values
            for key, values in result.items()
            if values
        }

        # ---------------------------------------------------------
        # VARIANTA 2
        # Material / Heat No.
        #
        # Např.:
        # AISI 4340 / 296137
        #
        # Rozdělujeme pouze na prvním "/".
        # ---------------------------------------------------------
        material_heat_number = self.find_values_next_to_labels(
            material_heat_number_labels
        )

        if material_heat_number:
            materials = []
            heat_numbers = []

            for value in material_heat_number:
                if "/" in value:
                    material, heat_number = value.split("/", 1)

                    materials.append(material.strip())
                    heat_numbers.append(heat_number.strip())
                else:
                    materials.append(value.strip())

            if materials:
                result["material"] = materials

            if heat_numbers:
                result["heat_number"] = heat_numbers

        job_number_note = self.find_values_next_to_labels(
            job_number_note_labels
        )

        if job_number_note:
            for value in job_number_note[:1]:
                if " / " in value:
                    job_number, note = value.split(" / ", 1)

                    result["job_number"] = [job_number.strip()]
                    result["note"] = [note.strip()]

        material_drawing = self.find_values_next_to_labels(
            material_drawing_labels
        )

        if material_drawing:
            for value in material_drawing[:1]:
                if " / " in value:
                    material, drawing_number = value.split(" / ", 1)

                    result["material"] = [material.strip()]
                    result["drawing_number"] = [drawing_number.strip()]

        job_number_drawing = self.find_values_next_to_labels(
            job_number_drawing_labels
        )
        

        if job_number_drawing:
            for value in job_number_drawing[:1]:
                if " / " in value:
                    job_number, drawing_number = value.split(" / ", 1)

                    result["job_number"] = [job_number.strip()]
                    result["drawing_number"] = [drawing_number.strip()]

        heat_number_material = self.find_values_next_to_labels(
            heat_number_material_labels
        )

       

        if heat_number_material:
            for value in heat_number_material[:1]:
                if " / " in value:
                    heat_number, material = value.split(" / ", 1)

                    result["heat_number"] = [heat_number.strip()]
                    result["material"] = [material.strip()]

        job_number_material = self.find_values_next_to_labels(
            job_number_material_labels
        )

        if job_number_material:
            for value in job_number_material[:1]:
                if " / " in value:
                    job_number, material = value.split(" / ", 1)

                    result["job_number"] = [job_number.strip()]
                    result["material"] = [material.strip()]

   
        # ---------------------------------------------------------
        # VARIANTA 3
        # Test code / Heat number
        #
        # Jednotlivé dvojice jsou oddělené ";"
        #
        # Např.:
        # J3/3 (84 pcs) / 397750/J3;
        # J17/1, 2 (121 pcs) / 397749/J17
        #
        # "/" uvnitř test code ani heat number nerozdělujeme.
        # Rozdělení probíhá podle " / ".
        # ---------------------------------------------------------
        test_code_heat_number = []

        values = self.find_values_next_to_labels(
            test_code_heat_number_labels
        )

        for value in values:
            for item in value.split(";"):
                item = item.strip()

                if not item:
                    continue

                if " / " in item:
                    test_code, heat_number = item.split(" / ", 1)

                    test_code_heat_number.append(
                        {
                            "test_code": test_code.strip(),
                            "heat_number": heat_number.strip(),
                        }
                    )
                else:
                    test_code_heat_number.append(
                        {
                            "test_code": item,
                        }
                    )
        if test_code_heat_number:
            result["test_code_heat_number"] = test_code_heat_number            

        return result    


    def get_cpf_customer_requirement(self) -> str | None:
        for paragraph in self.file.paragraphs:
            text = paragraph.text.strip()

            for label in customer_requirement_labels:
                if text.startswith(label):
                    value = text[len(label):].strip()

                    if value.startswith(":"):
                        value = value[1:].strip()

                    return value or None

        return None



    def get_cpf_tests(self) -> dict[str, list[dict[str, str]]]:
        for i, table in enumerate(self.tables):
            print("TABLE", i)
            print(table.table)

        tests = {}

        for name, (header_row, columns) in test_tables.items():

            # ---------------------------------------------------------
            # HARDNESS — beze změny
            # ---------------------------------------------------------
            if name == "hardness":
                hardness_table = None

                specimen_labels = ("Specimen", "Vzorek")
                hardness_labels = (
                    "Hardness measurements",
                    "Měření tvrdosti",
                )
                average_labels = ("Average", "Průměr")

                for table in self.tables:
                    header = list(table.table_lower[header_row])

                    if (
                        any(label.lower() in header for label in specimen_labels)
                        and any(label.lower() in header for label in hardness_labels)
                        and any(label.lower() in header for label in average_labels)
                    ):
                        hardness_table = table
                        break

                if hardness_table is None:
                    tests[name] = []
                    continue

                header = list(hardness_table.table_lower[header_row])

                specimen_idx = next(
                    i
                    for i, value in enumerate(header)
                    if value in [label.lower() for label in specimen_labels]
                )

                hardness_indices = [
                    i
                    for i, value in enumerate(header)
                    if value in [label.lower() for label in hardness_labels]
                ]

                average_idx = next(
                    i
                    for i, value in enumerate(header)
                    if value in [label.lower() for label in average_labels]
                )

                rows = []

                for row in hardness_table.table[header_row + 1:]:
                    if pd.isnull(row[specimen_idx]) or row[specimen_idx] == "":
                        continue

                    rows.append(
                        {
                            "Specimen": to_python_str(row[specimen_idx]),
                            "Hardness measurements": ", ".join(
                                to_python_str(row[i]).rstrip(",").strip()
                                for i in hardness_indices
                            ),
                            "Average": to_python_str(row[average_idx]),
                        }
                    )

                tests[name] = rows
                continue

            # ---------------------------------------------------------
            # IMPACT — EN / CZ / DE
            # ---------------------------------------------------------
            if name == "impact":
                impact_table = None
                impact_header_row = None

                column_aliases = {
                    "Specimen": (
                        "specimen",
                        "vzorek",
                        "probennummer",
                        "probenummer",
                    ),
                    "B": ("b",),
                    "W": ("w",),
                    "Ligament": ("ligament",),
                    "Temp.": ("temp.", "tepl.", "teplota"),
                    "KU2": ("ku2", "kv2"),
                    "Fracture": (
                        "fracture",
                        "rozlomení",
                        "fraktur",
                    ),
                }

                for table in self.tables:
                    for row_index in range(len(table.table_lower)):
                        header = list(table.table_lower[row_index])

                        if all(
                            any(alias.lower() in header for alias in aliases)
                            for aliases in column_aliases.values()
                        ):
                            impact_table = table
                            impact_header_row = row_index
                            break

                    if impact_table is not None:
                        break

                if impact_table is None:
                    tests[name] = []
                    continue

                header = list(
                    impact_table.table_lower[impact_header_row]
                )

                impact_indices = {
                    output_name: next(
                        i
                        for i, value in enumerate(header)
                        if value in [alias.lower() for alias in aliases]
                    )
                    for output_name, aliases in column_aliases.items()
                }

                rows = []

                for row in impact_table.table[impact_header_row + 2:]:
                    specimen = row[impact_indices["Specimen"]]

                    if pd.isnull(specimen) or str(specimen).strip() == "":
                        continue

                    if str(specimen).strip().lower() in (
                        "average",
                        "průměr",
                        "mittelwert",
                    ):
                        continue

                    rows.append(
                        {
                            output_name: to_python_str(row[index])
                            for output_name, index in impact_indices.items()
                        }
                    )

                tests[name] = rows
                continue

            # ---------------------------------------------------------
            # TENSILE — EN / CZ / DE
            # ---------------------------------------------------------
            if name == "tensile":
                tensile_table = None
                tensile_header_row = None

                column_aliases = {
                    "Specimen": (
                        "specimen",
                        "vzorek",
                        "probenummer",
                        "probennummer",
                    ),
                    "Temp.": ("temp.", "tepl.", "teplota"),
                    "d0": ("d0",),
                    "du": ("du",),
                    "l0": ("l0",),
                    "lu": ("lu",),
                    "rp0,2": ("rp0,2",),
                    "rm": ("rm",),
                    "ag": ("ag",),
                    "a": ("a",),
                    "z": ("z",),
                }

                for table in self.tables:
                    for row_index in range(len(table.table_lower)):
                        header = list(table.table_lower[row_index])

                        if all(
                            any(alias.lower() in header for alias in aliases)
                            for aliases in column_aliases.values()
                        ):
                            tensile_table = table
                            tensile_header_row = row_index
                            break

                    if tensile_table is not None:
                        break

                if tensile_table is None:
                    tests[name] = []
                    continue

                header = list(
                    tensile_table.table_lower[tensile_header_row]
                )

                tensile_indices = {
                    output_name: next(
                        i
                        for i, value in enumerate(header)
                        if value in [alias.lower() for alias in aliases]
                    )
                    for output_name, aliases in column_aliases.items()
                }

                rows = []

                for row in tensile_table.table[tensile_header_row + 2:]:
                    specimen = row[tensile_indices["Specimen"]]

                    if pd.isnull(specimen) or str(specimen).strip() == "":
                        continue

                    rows.append(
                        {
                            output_name: to_python_str(row[index])
                            for output_name, index in tensile_indices.items()
                        }
                    )

                tests[name] = rows

        return tests


    def get_cpf_methods(self) -> list[dict[str, str]]:
        methods = []

        for key, value in self.find_labelled_paragraphs(method_labels):
            if key == "name":
                method = {"name": value}

                for series, labels in test_series_labels.items():
                    if value.strip().lower() in [label.lower() for label in labels]:
                        method["test_series"] = series
                        break

                methods.append(method)

            elif methods:
                methods[-1][key] = value

        return methods

    def get_cpf_data(self) -> dict:
        data = {
            "source": self.name,
            "header": self.get_cpf_header(),
            "methods": self.get_cpf_methods(),
            "tests": self.get_cpf_tests(),
        }

        customer_requirement = self.get_cpf_customer_requirement()

        if customer_requirement:
            data["customer_requirement"] = customer_requirement

        return data

class CPFData(Data):
    excel_file_cls = ExcelFile
    word_file_cls = CPFWordFile

    def __init__(
        self,
        root: str,
        order_name: str,
        logger: logging.Logger,
    ) -> None:
        super().__init__(root, order_name, logger)
        self.cpf: list[dict] = []

    def save(self, path: str) -> None:
        import json

       

        with open(path, "w", encoding="utf-8") as f:
            json.dump(
                self.cpf,
                f,
                ensure_ascii=False,
                indent=2,
            )
    def use_file(self, file: ExcelFile | WordFile) -> bool:
        return (isinstance(file, CPFWordFile)
        )

    def extract_file_information(
        self,
        file: ExcelFile | WordFile,
    ) -> None:
        assert isinstance(file, CPFWordFile)
        self.cpf.append(file.get_cpf_data())