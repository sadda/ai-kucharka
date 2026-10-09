import argparse
import importlib
import io
import os
import platform
import sys
import tomllib
from typing import Any, cast

from tqdm import tqdm

from ai_kucharka import MSOffice, close_logger, filter_warnings, get_logger

DEFAULT_CONFIG = "configs/cpf.local.toml" 



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract information from order folders.")
    parser.add_argument("--config", default=DEFAULT_CONFIG, help=f"Path to the TOML config (default: {DEFAULT_CONFIG}).")
    parser.add_argument("orders", nargs="*", help="Orders to process (default: all orders in root_input).")
    return parser.parse_args()


def load_config(path: str) -> dict[str, Any]:
    with open(path, "rb") as f:
        return tomllib.load(f)


def load_class(name: str) -> type:
    module_name, class_name = name.rsplit(".", 1)
    return getattr(importlib.import_module(module_name), class_name)


args = parse_args()
config = load_config(args.config)
DataClass = load_class(config["data_class"])
root_input = config["root_input"]
root_output = config["root_output"]
auxiliary_folder = config["auxiliary_folder"]
orders_skip = config.get("orders_skip", [])
overwrite = config.get("overwrite", False)

stdout = cast(io.TextIOWrapper, sys.stdout)
stdout.reconfigure(encoding="utf-8")

filter_warnings()
logger = get_logger()
office = None
if platform.system() == "Windows":
    MSOffice.kill_all()
    office = MSOffice(logger)

os.makedirs(root_output, exist_ok=True)
order_names = args.orders if args.orders else sorted(os.listdir(root_input))

for order_name in tqdm(order_names):
    print(order_name)

    if order_name in orders_skip:
        continue

    root = f"{root_input}/{order_name}"

    data = DataClass(root, order_name, logger)
    data.extract_files(output_dir=auxiliary_folder, office=office)

    for file in data.files:
        if not data.use_file(file):
            continue

        data.cpf = []

        data.extract_file_information(file)

    

        header = data.cpf[0]["header"]
        order_number = header.get("order_number", [order_name])[0]

        language = os.path.splitext(file.name)[0].split("_")[-1].upper()

        result_name = (
            f"{root_output}/"
            f"{order_number.replace('/', '_')}_{language}.json"
        )

        if overwrite or not os.path.exists(result_name):
            data.save(result_name)

    data.close_files()
close_logger()
