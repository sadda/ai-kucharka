import io
import os
import platform
import sys
from typing import cast

import numpy as np
from tqdm import tqdm

from processing import MSOffice, close_logger, filter_warnings, get_logger
from processing.company123.data import Company123Data as DataClass

stdout = cast(io.TextIOWrapper, sys.stdout)
stdout.reconfigure(encoding="utf-8")

filter_warnings()
logger = get_logger()
office = None
if platform.system() == "Windows":
    MSOffice.kill_all()
    office = MSOffice(logger)

root_input = "data/Company123"
root_output = "results_company123"
auxiliary_folder = "auxiliary"
orders_skip = [".DS_Store", "CN1311055", "ZAK1301031"]

os.makedirs(root_output, exist_ok=True)
order_names = np.array(os.listdir(root_input))

for order_name in tqdm(order_names):
    print(order_name)
    if order_name in orders_skip:
        continue
    root = f"{root_input}/{order_name}"
    result_name = f"{root_output}/{order_name}.pkl"
    if os.path.exists(result_name):
        data = DataClass.load(result_name, logger)
    else:
        data = DataClass(root, order_name, logger)
        data.extract_files(output_dir=auxiliary_folder, office=office)
        data.extract_information()
        data.save(result_name)
    data.close_files()
close_logger()
