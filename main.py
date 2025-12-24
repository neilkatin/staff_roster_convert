#! /usr/bin/env python
# main.py - convert staff services staff rosters into a more useful form

import os
import re
import sys
import logging
import argparse
import datetime
import io
import typing

import xlrd
import dotenv

import openpyxl
import openpyxl.utils
import openpyxl.styles
import openpyxl.styles.colors
import openpyxl.writer.excel

import config as config_static
import neil_tools
from neil_tools import spreadsheet_tools


# index of column label row, origin zero
STAFF_ROSTER_LABEL_ROW = 5



def main() -> None:
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    errors = False
    input_file_name = "Staff Roster_Dec 23, 2025 9_00_00 AM.xls"
    input_file = read_xls_file(input_file_name)

    if errors:
        sys.exit(1)


def row_fixups(label_row):

    right_aligned = openpyxl.styles.Alignment(horizontal="right")

    fixup_defs = {
            'Name': { 'width': 25, },
            'Preferred name': { 'width': 10 },
            'Region': { 'width': 6, },
            'State': { 'width': 4, },
            'Res': { 'width': 4, },
            'T&M': { 'width': 4, },
            'GAP(s)': { 'width': 15, },
            'District': { 'width': 5, },
            'Qualification (assignment)': { 'width': 5, },
            'Current/Last Supervisor': { 'width': 30, },
            'Reporting/Work Location': { 'width': 30, },
            'On Job': { 'width': 5, },
            'DaysRemain': { 'width': 5, },
            '# dep': { 'width': 5, },
            'Lodging Last Night': { 'width': 30, },
            'Lodging Tonight': { 'width': 30, },
            'Qualifications (member)': { 'width': 30, },
            'All GAPs': { 'width': 30, },
            'All Supervisors': { 'width': 30, },
            'Work Location': { 'width': 30, },
            'Email': { 'width': 30, },

            'Assigned': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                         'number_format': "yyyy-mm-dd",
                         },
            'Checked in': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                           'number_format': "yyyy-mm-dd",
                         },
            'Released': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                         'number_format': "yyyy-mm-dd",
                         },
            'Travel home': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                         'number_format': "yyyy-mm-dd",
                         },
            'Last Daily Checkin': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                         'number_format': "yyyy-mm-dd",
                         },
            'DaysRemain': { 'width': 5, 'number_format': "##0",
                           'alignment': right_aligned,
                           'convert_value': lambda x: x if isinstance(x, int) else x if  x == '' or x == 'n/a' else int(x),
                           },
        }

    fixups_by_col = {}

    for c in range(0, len(label_row)):
        name = label_row[c]

        if name in fixup_defs:
            fixups_by_col[c] = fixup_defs[name]
        else:
            fixups_by_col[c] = {}

    return fixups_by_col

def fixup_cell_header(ws, c, fixup):

    col_letter = openpyxl.utils.get_column_letter(c +1)

    if 'width' in fixup:
        #log.debug(f"setting col { c } to { fixup['width'] }")
        ws.column_dimensions[col_letter].width = fixup['width']
    else:
        ws.column_dimensions[col_letter].auto_size = True


def fixup_cell(cell, fixup):

    if 'convert_value' in fixup:
        log.debug(f"cell { cell } old value { cell.value } isint { isinstance(cell.value, int) }")
        cell.value = fixup['convert_value'](cell.value)
    if 'number_format' in fixup:
        cell.number_format = fixup['number_format']
    if 'alignment' in fixup:
        log.debug(f"setting cell { cell } alignment { fixup['alignment'] }")
        cell.alignment = fixup['alignment']




def read_xls_file(filename: str) -> None:
    
    book_in = xlrd.open_workbook(filename)
    sheet_in = book_in.sheet_by_index(0)

    log.debug(f"sheet name { sheet_in.name } rows { sheet_in.nrows } cols { sheet_in.ncols }")

    label_row = STAFF_ROSTER_LABEL_ROW

    #for col in range(0, sheet.ncols):
    #    cell_value = sheet.cell_value(label_row, col)
    #    log.debug(f"cell({ label_row }, { col }) = { cell_value }")

    label_values = sheet_in.row_values(label_row)
    log.debug(f"label_values: { label_values }")


    # copy everything to a clean workbook
    book_out = openpyxl.Workbook()
    sheet_out = book_out.active
    sheet_out.title = "Roster"

    # set column attributes
    fixups_by_col = row_fixups(label_values)
    for c in range(0, len(label_values)):
        fixup_cell_header(sheet_out, c, fixups_by_col[c])
        

    # copy cells
    for r in range(0, sheet_in.nrows):
        row = sheet_in.row_values(r)

        for c in range(0, sheet_in.ncols):
            value = row[c]
            cell = sheet_out.cell(row=r +1, column=c +1, value=value)
            # don't fix up cells before the actual data
            if r > label_row:
                fixup_cell(cell, fixups_by_col[c])

    # make a table if there is data
    if sheet_in.nrows > label_row + 1:
        last_col_letter = openpyxl.utils.get_column_letter(sheet_in.ncols)
        table_ref = f"A{label_row + 1}:{ last_col_letter }{ sheet_in.nrows - label_row }"
        log.debug(f"adding table; table_ref '{ table_ref }'")
        table = openpyxl.worksheet.table.Table(displayName="Roster", ref=table_ref)
        sheet_out.add_table(table)

        sheet_out.freeze_panes = f"B{ label_row + 2 }"
        log.debug(f"freeze_panes: { sheet_out.freeze_panes }")

    book_out.save("test.xlsx")


    
    




def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
            description="tool to convert staffing reports into a more usefull form",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")
    parser.add_argument("--out", help="file to save output into", action="store_true")

    args = parser.parse_args()

    return args


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)
