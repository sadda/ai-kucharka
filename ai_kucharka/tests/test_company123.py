import logging
import os

import pytest

from ai_kucharka.company123.config import ATTRIBUTES
from ai_kucharka.company123.data import Company123Data

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
ROOT_INPUT = os.path.join(ROOT, "data", "Company123")

EXPECTED = {
    "ZAK001": {
        "Číslo zakázky": "ZAK-2026-001",
        "Zákazník": "MechaWorks s.r.o.",
        "Výrobek": "Nosník",
        "Materiál": "S355",
        "Množství": "17",
        "Termín dodání": "04.10.2026",
        "Priorita": "Normální",
        "Stav": "Ve výrobě",
        "Cena": "26 375 Kč",
        "Projektový manažer": "Jan Novák",
    },
    "ZAK002": {
        "Číslo zakázky": "ZAK-2026-002",
        "Zákazník": "TechMetal s.r.o.",
        "Výrobek": "Hřídel",
        "Materiál": "C45",
        "Množství": "24",
        "Termín dodání": "07.11.2026",
        "Priorita": "Vysoká",
        "Stav": "Dokončená",
        "Cena": "27 750 Kč",
        "Projektový manažer": "Petr Svoboda",
    },
}


def extract(order_name: str, output_dir: str) -> Company123Data:
    data = Company123Data(os.path.join(ROOT_INPUT, order_name), order_name, logging.getLogger(__name__))
    data.extract_files(output_dir=output_dir)
    data.extract_information()
    return data


@pytest.mark.parametrize("order_name", sorted(EXPECTED))
def test_extract_information(order_name: str, tmp_path) -> None:
    data = extract(order_name, str(tmp_path))
    assert len(data.files) == 1
    assert set(data.info) == set(ATTRIBUTES)
    for attribute, value in EXPECTED[order_name].items():
        assert [x.value for x in data.info[attribute]] == [value]
        assert all(os.path.basename(x.source).startswith("Zakazka2026") for x in data.info[attribute])
    data.close_files()


@pytest.mark.parametrize("order_name", sorted(EXPECTED))
def test_pickle_roundtrip(order_name: str, tmp_path) -> None:
    data = extract(order_name, str(tmp_path))
    path = tmp_path / f"{order_name}.pkl"
    data.save(str(path))
    data.close_files()

    loaded = Company123Data.load(str(path))
    assert loaded.order_name == order_name
    assert loaded.file_names == data.file_names
    assert loaded.files == []
    assert loaded.info == data.info
