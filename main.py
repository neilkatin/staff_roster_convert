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
import base64

import xlrd
import dotenv

import openpyxl
import openpyxl.utils
import openpyxl.styles
import openpyxl.styles.colors
import openpyxl.writer.excel
import O365

import config as config_static
import neil_tools
import arc_o365
from neil_tools import spreadsheet_tools


# index of column label row, origin zero
STAFF_ROSTER_LABEL_ROW = 5
ORIG_SHEET_NAME = "Orig"



def main() -> None:
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    report_dict = fetch_workforce_reports(config, "033-26", config.TOKEN_FILENAME_AVIS)

    errors = False
    book_out = openpyxl.Workbook()

    # do the 'orig' roster first so it is at the end of the list
    sheet_orig = read_roster(book_out, ORIG_SHEET_NAME, report_dict['Staff Roster'], STAFF_ROSTER_LABEL_ROW, ROSTER_FIXUPS)

    # now copy and filter the 'orig' sheet to the others
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Active", filter_row_active)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "MC", filter_row_mc)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "CC", filter_row_cc)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "LOG", filter_row_log)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "ER", filter_row_er)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "IP", filter_row_ip)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "WF", filter_row_wf)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "OM", filter_row_om)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Overstayed", filter_row_overstayed)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Days_2", filter_row_2_days)
    #copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Needs_Sup", filter_row_needs_sup)


    read_roster(book_out, 'Staff Requests', report_dict['Open Staff Requests'], 1, ROSTER_FIXUPS)
    read_roster(book_out, 'Shifts', report_dict['DRO Shift Tool - Shift Registrant Details'], 3, ROSTER_FIXUPS)
    read_roster(book_out, 'Air', report_dict['Air Travel Roster'], 2, ROSTER_FIXUPS, freeze_col="C")
    read_roster(book_out, 'Arrival', report_dict['Arrival Roster'], 5, ROSTER_FIXUPS, suppress_columns={'Z':True})

    # now copy and filter the 'orig' sheet to the others
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Needs_Sup", filter_row_needs_sup, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Days_2", filter_row_2_days, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Overstayed", filter_row_overstayed, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "OM", filter_row_om, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "WF", filter_row_wf, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "IP", filter_row_ip, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "ER", filter_row_er, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "LOG", filter_row_log, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "CC", filter_row_cc, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "MC", filter_row_mc, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Roster", filter_row_active, ROSTER_FIXUPS)

    del book_out['Sheet']

    if errors:
        sys.exit(1)

    book_out.save("test.xlsx")



MATCH_TO_FIRST_UNDERSCORE = re.compile(r"^([^_]+)_")
def fetch_workforce_reports(config, dro_id, token_filename):

    account = init_o365(config, token_filename)

    message_match_string = f"DR { dro_id } Automated Workforce Reports"
    message = search_mail(account, config.PROGRAM_EMAIL, message_match_string)

    if message is None:
        error = f"Could not find an email that matches '{ message_match_string }'"
        log.fatal(error)
        raise(Exception(error))

    attach_dict = {}
    # read the attachments
    for attachment in message.attachments:
        attach_name = attachment.name
        name_type = attach_name
        name_before_underscore = MATCH_TO_FIRST_UNDERSCORE.match(attach_name)

        if name_before_underscore is not None:
            name_type  = name_before_underscore.group(1)

        log.debug(f"attachment { attachment.name } name_type { name_type } size { attachment.size }")
        attach_dict[name_type] = base64.b64decode(attachment.content)

    return attach_dict





def search_mail(account, email_address, subj_pattern):

    mailbox = account.mailbox(resource=email_address)
    builder = mailbox.new_query()
    dt = datetime.datetime(1900, 1, 1)
    matcher = builder.chain_and(
            builder.greater('sentDateTime', dt),
            builder.contains("subject", subj_pattern))

    messages = mailbox.get_messages(query=matcher, order_by="sentDateTime desc", limit=1, download_attachments=True)
    message = next(messages, None)

    if message is not None:
        log.debug(f"returning message { message.subject } received { message.received }")
    else:
        log.debug(f"No message found")

    return message




def init_o365(config, token_filename):
    """ do initial setup to get a handle on office 365 graph api """

    o365 = arc_o365.arc_o365.arc_o365(config, token_filename=token_filename, timezone="America/Los_Angeles")

    account = o365.get_account()
    if account is None:
        raise Exception("could not access office 365 graph api")

    return account


RIGHT_ALIGNED = openpyxl.styles.Alignment(horizontal="right")
ROSTER_FIXUPS = {
        'Name': { 'width': 25, },
        'Preferred name': { 'width': 10 },
        'Region': { 'width': 6, },
        'State': { 'width': 4, },
        'Res': { 'width': 4, },
        'T&M': { 'width': 6, },
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
                       'alignment': RIGHT_ALIGNED,
                       'convert_value': lambda x: x if isinstance(x, int) else x if  x == '' or x == 'n/a' else int(x),
                       },
    }


def row_fixups(fixup_defs, label_row):


    fixups_by_col = {}

    # generate a mapping from column name to column index
    column_name_map = {}

    for c in range(0, len(label_row)):
        name = label_row[c]

        if name in fixup_defs:
            fixups_by_col[c] = fixup_defs[name]
        else:
            fixups_by_col[c] = {}

        column_name_map[name] = c


    return fixups_by_col, column_name_map




def fixup_cell_header(ws, c, fixup):

    col_letter = openpyxl.utils.get_column_letter(c +1)

    if 'width' in fixup:
        #log.debug(f"setting col { c } to { fixup['width'] }")
        ws.column_dimensions[col_letter].width = fixup['width']
    else:
        ws.column_dimensions[col_letter].auto_size = True


def fixup_cell(cell, fixup):

    if 'convert_value' in fixup:
        #log.debug(f"cell { cell } old value { cell.value } isint { isinstance(cell.value, int) }")
        cell.value = fixup['convert_value'](cell.value)
    if 'number_format' in fixup:
        cell.number_format = fixup['number_format']
    if 'alignment' in fixup:
        #log.debug(f"setting cell { cell } alignment { fixup['alignment'] }")
        cell.alignment = fixup['alignment']


filter_row_active = { 'Released': lambda x: x == '' }
filter_row_overstayed = { 'DaysRemain': lambda x: x != '' and x != 'n/a' and x is not None and int(x) < 0 }
filter_row_2_days = { 'DaysRemain': lambda x: x != '' and x != 'n/a' and x is not None and int(x) == 2 }
filter_row_needs_sup = { 'Current/Last Supervisor': lambda x: x != '' and x is not None and x == 'Needs Supervisor' }
filter_row_mc = { 'GAP(s)': lambda x: isinstance(x, str) and x.startswith('MC/') }
filter_row_cc = { 'GAP(s)': lambda x: isinstance(x, str) and x.startswith('CC/') }
filter_row_log = { 'GAP(s)': lambda x: isinstance(x, str) and x.startswith('LOG/') }
filter_row_er = { 'GAP(s)': lambda x: isinstance(x, str) and x.startswith('ER/') }
filter_row_ip = { 'GAP(s)': lambda x: isinstance(x, str) and x.startswith('IP/') }
filter_row_wf = { 'GAP(s)': lambda x: isinstance(x, str) and x.startswith('WF/') }
filter_row_om = { 'GAP(s)': lambda x: isinstance(x, str) and x.startswith('OM/') }

def filter_row(row, row_name_map, filter_defs):
    # check all the conditions in the filter; if they all pass include the row

    for name, func in filter_defs.items():
        #log.debug(f"filter_row: name '{ name }'")
        c = row_name_map[name]
        val = row[c]
        #log.debug(f"filter_row: name { name } c { c } val '{ val }' ")
        if not func(val):
            #log.debug(f"filter_row: name { name } val '{ val }' returning False")
            return False
        
    #log.debug(f"filter_row: returning True")
    return True


def read_roster(book_out, sheet_name, file_contents: str, label_row: int, fixups: dict, freeze_col: str = "B", suppress_columns: dict[str] = {}) -> openpyxl.worksheet.worksheet:
    
    book_in = xlrd.open_workbook(file_contents=file_contents)
    sheet_in = book_in.sheet_by_index(0)

    log.debug(f"sheet name { sheet_in.name } rows { sheet_in.nrows } cols { sheet_in.ncols }")

    #for col in range(0, sheet.ncols):
    #    cell_value = sheet.cell_value(label_row, col)
    #    log.debug(f"cell({ label_row }, { col }) = { cell_value }")

    label_values = sheet_in.row_values(label_row)
    #log.debug(f"label_values: { label_values }")


    # copy everything to a clean workbook
    sheet_orig = book_out.create_sheet(sheet_name, 0)
    sheet_orig.title = sheet_name

    # set column attributes
    fixups_by_col, column_name_map = row_fixups(fixups, label_values)
    for c in range(0, len(label_values)):
        fixup_cell_header(sheet_orig, c, fixups_by_col[c])
        

    # copy cells
    for r in range(0, sheet_in.nrows):
        row = sheet_in.row_values(r)


        output_c = 0
        for c in range(0, sheet_in.ncols):
            col_letter = openpyxl.utils.get_column_letter(c +1)
            if col_letter in suppress_columns:
                # skip this column
                #log.debug(f"read_roster: suppressing columm '{ col_letter }'")
                continue

            
            value = row[c]
            cell = sheet_orig.cell(row=r +1, column=output_c +1, value=value)

            # don't fix up cells before the actual data
            if r > label_row:
                fixup_cell(cell, fixups_by_col[c])

            output_c = output_c + 1


    # make a table if there is data
    if sheet_in.nrows > label_row + 1:
        last_col_letter = openpyxl.utils.get_column_letter(sheet_orig.max_column)
        table_ref = f"A{label_row + 1}:{ last_col_letter }{ sheet_in.nrows }"
        log.debug(f"adding table; table '{ sheet_name }' table_ref '{ table_ref }'")
        table = openpyxl.worksheet.table.Table(displayName=sheet_name, ref=table_ref)
        sheet_orig.add_table(table)

        sheet_orig.freeze_panes = f"{ freeze_col}{ label_row + 2 }"

    return sheet_orig


# copy from the 'orig' sheet to a new sheet, filtering entries
def copy_sheet(wb, sheet_orig, label_row, sheet_name, filters, fixups):
    
    log.debug(f"copy_sheet: sheet_name { sheet_name } label_row { label_row }")
    #sheet_new = wb.create_sheet(sheet_name, len(wb.sheetnames)-1)
    sheet_new = wb.create_sheet(sheet_name, 0)

    label_values = list(next(sheet_orig.iter_rows(min_row=label_row +1, max_row=label_row +2, values_only=True)))

    # set column attributes
    fixups_by_col, column_name_map = row_fixups(fixups, label_values)
    for c in range(0, len(label_values)):
        fixup_cell_header(sheet_new, c, fixups_by_col[c])

    # these two are origin one indexes
    max_col = sheet_orig.max_column
    max_row = sheet_orig.max_row

    # copy cells
    output_row = 1
    for r in range(label_row, max_row):
        row = list(sheet_orig.iter_rows(min_row=r +1, max_row=r +2, values_only=True))
        row_values = list(row[0])
        #log.debug(f"copy_sheet: row { row }")

        include_row = False
        if r > label_row:
            include_row = filter_row(row_values, column_name_map, filters)
        else:
            include_row = True

        if not include_row:
            #log.debug(f"row { r } output_row { output_row } not included")
            continue

        for c in range(0, max_col):
            cell = sheet_new.cell(row=output_row, column=c +1, value=row_values[c])

            # don't fix up cells before the actual data
            if r > label_row:
                fixup_cell(cell, fixups_by_col[c])

        # increment the output row
        output_row = output_row + 1


    # make a table if there is data
    if output_row > 2:
        last_col_letter = openpyxl.utils.get_column_letter(max_col)

        table_ref = f"A1:{ last_col_letter }{ output_row -1 }"
        log.debug(f"table { sheet_name } table_ref '{ table_ref }'")
        table_new = openpyxl.worksheet.table.Table(displayName=sheet_name, ref=table_ref)
        sheet_new.add_table(table_new)
        sheet_new.freeze_panes = f"B2"



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
