import gzip
import json
import time
import os
from datetime import datetime
from typing import List
import concurrent.futures

import xlsxwriter


class Record:
    def __init__(self, data):
        self.id = data.get("id")
        self.my_string_1 = data.get("myString1")
        self.my_date_1 = data.get("myDate1")
        self.my_date_2 = data.get("myDate2")
        self.amount = data.get("amount")
        self.my_numeric_string2 = data.get("myNumericString")
        self.my_string_2 = data.get("myString2")


def get_content() -> List[Record]:
    with gzip.open("../input.json.gzip", "rt", encoding="utf-8") as f:
        content = f.read()
    
    data = json.loads(content)
    records = [Record(item) for item in data]
    return records


def write_sheet(worksheet, records, workbook):
    decimal_format = workbook.add_format({"num_format": "0.000"})
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    worksheet.set_column(0, 0, 22)

    for i, rec in enumerate(records):
        worksheet.write(i, 0, rec.id)
        worksheet.write(i, 1, rec.my_string_1)
        worksheet.write(i, 2, rec.my_numeric_string2 or "")
        worksheet.write(i, 3, rec.my_string_2 or "")
        worksheet.write(i, 4, rec.amount, decimal_format)

        my_date_2 = datetime.strptime(rec.my_date_2, "%Y-%m-%d")
        worksheet.write_datetime(i, 5, my_date_2, date_format)

        my_date_1 = datetime.strptime(rec.my_date_1, "%Y-%m-%d")
        worksheet.write_datetime(i, 6, my_date_1, date_format)


def to_excel(records: List[Record]):
    n_sheets = int(os.environ.get('N_SHEETS', '1'))
    n_sheets = max(1, min(n_sheets, 9))  # Ensure n_sheets is between 1 and 9

    workbook = xlsxwriter.Workbook("demo.xlsx", {'constant_memory': True})
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=n_sheets) as executor:
        futures = []
        for i in range(n_sheets):
            worksheet = workbook.add_worksheet(f"Sheet{i+1}")
            futures.append(executor.submit(write_sheet, worksheet, records, workbook))
        
        # Wait for all sheets to be written
        concurrent.futures.wait(futures)

    workbook.close()


def main():
    start = time.time()
    records = get_content()
    print(f"Load Time {time.time() - start:.2f} seconds")

    start = time.time()
    to_excel(records)
    print(f"Xlsx Write Time {time.time() - start:.2f} seconds")


if __name__ == "__main__":
    main()