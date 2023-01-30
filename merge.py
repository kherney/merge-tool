import pandas as pd
from pandas.io.excel import ExcelFile
from pandas.core.indexes.base import Index
from pandas import isna

from argparse import ArgumentParser

from typing import Dict, List, Any


def update_data(data: Dict, cell: Any, key: Any, n_cell: int, delimiter: str = ','):
    on_memory = data.get(key)
    cell_value = on_memory[n_cell]
    cell_value = '' if cell_value is None else cell_value

    if str(cell) in cell_value.split(delimiter):
        return

    on_memory[n_cell] = str(cell) if not cell_value else cell_value + delimiter + str(cell)
    data.update({key: on_memory})


def main():
    data = {}

    def create_row():
        if key not in data.keys():
            data.update({key: ['' for _ in range(len(row))]})

    n_column = args.merge_by
    file = ExcelFile(args.in_file, engine='openpyxl')
    dt = pd.read_excel(file)
    labels = dt.columns.values
    data.update({'label': labels})

    for index, series in dt.iterrows():
        row = series.values
        key = row[n_column]
        create_row()
        for n_cell, cell in enumerate(row):
            if pd.isna(cell):
                continue
            update_data(data, cell, key, n_cell)

    df = pd.DataFrame(data=data.values())

    file_name = args.out_file
    if not file_name:
        file_name = "File Merged"
    df.to_excel("{}.xlsx".format(file_name), sheet_name='Sheet_name_1')


if __name__ == '__main__':

    parser = ArgumentParser(prog="Merge Tool",
                            conflict_handler='resolve',
                            description=" Script for merge rows by column name")
    parser.add_argument("--merge_by", type=int, dest="merge_by", required=True,
                        help="N: Positional Column in the given file. From zero 0 to N-1")
    parser.add_argument("--file", type=str, dest="in_file", required=True,
                        help="File to load")
    parser.add_argument("-o", "--out_path", type=str, dest="out_file", required=False,
                        help="File to output")
    args = parser.parse_args()

    main()
