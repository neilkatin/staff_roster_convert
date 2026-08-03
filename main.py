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
REPORT_DATE_FORMAT = "%Y-%m-%d %H-%M-%S %Z"


NOW = datetime.datetime.now().astimezone()


def main() -> None:
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    # read static configuration
    config = neil_tools.init_config(config_static, ".env")

    # make sure we know about this DRO
    dr_config = config.DRConfig.lookup_dr(args.dr_id)
    if dr_config is None:
        log.fatal(f"no configuration on file for '{ args.dr_id }'")
        sys.exit(1)

    o365 = arc_o365.arc_o365.arc_o365(config, token_filename=config.TOKEN_FILENAME, timezone="America/Los_Angeles")
    report_dict = o365.fetch_workforce_reports(dr_config.dr_id)
    report_date = report_dict['created']
    report_date_stamp = report_date.strftime(REPORT_DATE_FORMAT)
    log.debug(f"report date is '{ report_date }', stamp '{ report_date_stamp }'")

    errors = False
    book_out = openpyxl.Workbook()

    # do the 'orig' roster first so it is at the end of the list
    sheet_orig = read_roster(book_out, ORIG_SHEET_NAME, report_dict['Staff Roster - Cumulative'], STAFF_ROSTER_LABEL_ROW, ROSTER_FIXUPS)
    sheet_roster = copy_sheet(book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Roster", filter_row_active, ROSTER_FIXUPS, suppress_columns={'I': True})

    read_roster(book_out, 'StaffRequests', report_dict['Open Staff Requests'], 1, ROSTER_FIXUPS)
    read_roster(book_out, 'Shifts', report_dict['DRO Shift Tool - Shift Registrant Details'], 3, SHIFTS_FIXUPS)
    read_roster(book_out, 'Air', report_dict['Air Travel Roster'], 2, AIR_FIXUPS, freeze_col="C", suppress_columns={'V':True})
    read_roster(book_out, 'Arrival', report_dict['Arrival Roster'], 5, ARRIVAL_FIXUPS, suppress_columns={'Z':True})

    # now copy and filter the 'orig' sheet to the others
    copy_sheet(book_out, sheet_roster, 0, "Need_SMS", filter_row_sms, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "Needs_Sup", filter_row_needs_sup, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "Days_2", filter_row_2_days, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "Overstayed", filter_row_overstayed, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "OM", filter_row_om, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "WF", filter_row_wf, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "IP", filter_row_ip, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "ER", filter_row_er, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "LOG", filter_row_log, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "CC", filter_row_cc, ROSTER_FIXUPS)
    copy_sheet(book_out, sheet_roster, 0, "MC", filter_row_mc, ROSTER_FIXUPS)

    del book_out['Sheet']

    # move reoster to beginning of workbook
    # we should be using workbook.move_sheet(), but that didn't seem to work
    book_out.remove(sheet_roster)
    book_out._add_sheet(sheet_roster, index=0)

    if errors:
        sys.exit(1)

    roster_file_name = f"DR{ dr_config.dr_id } Staffing Report { report_date_stamp }.xlsx"

    book_out.save(roster_file_name)

    if args.send or args.test_send:
        send_roster(dr_config, args, o365.account, roster_file_name, report_date)

    if not args.save:
        os.remove(roster_file_name)



RIGHT_ALIGNED = openpyxl.styles.Alignment(horizontal="right")
ROSTER_FIXUPS = {
        'Name': { 'width': 25, },
        'Preferred name': { 'width': 10 },
        'Region': { 'width': 6, },
        'State': { 'width': 4, },
        'Res': { 'width': 4, },
        'T&M': { 'width': 6, },
        'GAP(s)': { 'width': 15, },
        'G/A/P': { 'width': 15, },
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
        'On Job': { 'width': 4, 'convert_value': lambda x: int(x) if isinstance(x, str) else x },
    }

ARRIVAL_FIXUPS = {
        'Region': { 'width': 6, },
        'Status': { 'width': 8, },
        'Resp': { 'width': 6, },
        'Category': { 'width': 6, },
        'Category': { 'width': 6, },
        'Gender': { 'width': 10, },
        'T&M': { 'width': 8, },
        'Trans': { 'width': 8, },
        'Flight Arrival Date/Time': { 'width': 8, },
        'Type': { 'width': 8, },
        'GAP': { 'width': 15, },
        '# Deploy': { 'width': 15, },
        'Texts?': { 'width': 8, },
        'Arrive date': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd",
                     },
        'Flight Arrival Date/Time': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        }



AIR_FIXUPS = {
        'Departure City': { 'width': 16 },
        'Arrival City': { 'width': 16 },
        'Ticketed': { 'width': 4 },
        'Airline': { 'width': 15 },
        'Flight': { 'width': 8 },
        'Region name': { 'width': 24 },
        'Reporting or Work location': { 'width': 24 },
        'District': { 'width': 10 },

        'Last action date': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'Exp Arrival': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd",
                     },
        'Departure time': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'Arrival time': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        }


SHIFTS_FIXUPS = {
        'Shift Name': { 'width': 30 },
        'County': { 'width': 24 },
        'Registration Status': { 'width': 16 },
        'Current Volunteer Status': { 'width': 20 },
        'District (of shift)': { 'width': 20 },
        'County (residence)': { 'width': 24 },
        'Registration Comments': { 'width': 24 },
        'Name': { 'width': 24 },
        'Registration Status': { 'width': 24 },
        'Ever DEBV/P-DEBV for this DRO': { 'width': 16 },
        'Email': { 'width': 24 },

        'Start Date': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd",
                     },
        'Start Time': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'End Date': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'Date Registered/Last Changed': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'End Time': { 'convert_value': lambda c: '' if c == '' else spreadsheet_tools.excel_to_dt(c),
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        }




def row_fixups(fixup_defs, label_row, supress_columns):


    fixups_by_col = {}

    # generate a mapping from column name to column index
    column_name_map = {}
    output_column = 0

    for c in range(0, len(label_row)):
        name = label_row[c]

        col_letter = openpyxl.utils.get_column_letter(c +1)
        if col_letter in supress_columns:
            # skip this column
            #log.debug(f"read_roster: suppressing columm '{ col_letter }' c { c } output_column { output_column }")
            continue


        if name in fixup_defs:
            fixups_by_col[output_column] = fixup_defs[name]
        else:
            fixups_by_col[output_column] = {}

        column_name_map[name] = output_column
        output_column += 1


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
filter_row_sms = { 'Texts?': lambda x: x != 'opt-in' }
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

        # ignore filtered columns
        if name in row_name_map:
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
    fixups_by_col, column_name_map = row_fixups(fixups, label_values, suppress_columns)
    output_col = 0
    for c in range(0, len(label_values)):
        col_letter = openpyxl.utils.get_column_letter(c +1)
        if col_letter in suppress_columns:
            continue

        if output_col not in fixups_by_col:
            log.error(f"read_roster: missing key: c { c } col_letter '{ col_letter }' output_col { output_col }")
        fixup_cell_header(sheet_orig, output_col, fixups_by_col[output_col])
        output_col += 1
        

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
                fixup_cell(cell, fixups_by_col[output_c])

            output_c = output_c + 1


    # make a table if there is data
    if sheet_orig.max_row > label_row:
        last_col_letter = openpyxl.utils.get_column_letter(sheet_orig.max_column)
        table_ref = f"A{label_row + 1}:{ last_col_letter }{ sheet_orig.max_row }"
        log.debug(f"adding table: table '{ sheet_name }' table_ref '{ table_ref }'")
        table = openpyxl.worksheet.table.Table(displayName=sheet_name, ref=table_ref)
        sheet_orig.add_table(table)

        sheet_orig.freeze_panes = f"{ freeze_col}{ label_row + 2 }"

    return sheet_orig


# copy from the 'orig' sheet to a new sheet, filtering entries
def copy_sheet(wb, sheet_orig, label_row, sheet_name, filters, fixups, suppress_columns: dict[str] = {}):
    
    #log.debug(f"copy_sheet: sheet_name { sheet_name } label_row { label_row }")
    #sheet_new = wb.create_sheet(sheet_name, len(wb.sheetnames)-1)
    sheet_new = wb.create_sheet(sheet_name, 0)

    label_values = list(next(sheet_orig.iter_rows(min_row=label_row +1, max_row=label_row +2, values_only=True)))

    # set column attributes
    fixups_by_col, column_name_map = row_fixups(fixups, label_values, suppress_columns)
    output_col = 0
    for c in range(0, len(label_values)):
        col_letter = openpyxl.utils.get_column_letter(c +1)
        if col_letter in suppress_columns:
            continue

        if output_col not in fixups_by_col:
            log.error(f"copy_sheet: c not in fixups_by col.  c { c } output_col { output_col } col_letter { col_letter }")
        fixup_cell_header(sheet_new, output_col, fixups_by_col[output_col])
        output_col += 1

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

        output_col = 0
        for c in range(0, max_col):
            col_letter = openpyxl.utils.get_column_letter(c+1)
            if col_letter in suppress_columns:
                # skip this column
                #log.debug(f"copy_sheet: suppressing columm '{ col_letter }'")
                continue

            cell = sheet_new.cell(row=output_row, column=output_col +1, value=row_values[c])

            # don't fix up cells before the actual data
            if r > label_row:
                fixup_cell(cell, fixups_by_col[output_col])
            output_col += 1

        # increment the output row
        output_row = output_row + 1


    # make a table if there is data
    if sheet_new.max_row > 2:
        last_col_letter = openpyxl.utils.get_column_letter(sheet_new.max_column)

        table_ref = f"A1:{ last_col_letter }{ sheet_new.max_row }"
        log.debug(f"copy_sheet: adding table { sheet_name } table_ref '{ table_ref }'")
        table_new = openpyxl.worksheet.table.Table(displayName=sheet_name, ref=table_ref)
        sheet_new.add_table(table_new)
        sheet_new.freeze_panes = f"B2"

    return sheet_new


#
# wrapper for sending out the roster
#
def send_roster(dr_config, args, account, file_name, report_date):

    warn_days = 2
    if report_date < NOW - datetime.timedelta(days=warn_days):
        date_warning = f"<span style='background-color:yellow;'>WARNING: staff report is more than { warn_days } days old<span>."
    else:
        date_warning = ""

    message_body = \
f"""
<p>
This report is based on the staff roster dated { report_date.strftime(REPORT_DATE_FORMAT) }.
</p>
{ date_warning }

Hello everyone.  This is a spreadsheet based on the automated staffing reports, but reorganized to hopefully be easier to use.
"""

    send_report_common(dr_config, args, account, file_name, "Staffing Report", message_body)


def send_report_common(dr_config, args, account, file_name, report_type, message_body):

    #message = account.new_message(resource=dr_config.send_email)
    message = account.new_message()

    if args.test_send:
        message.bcc.add(dr_config.to_test)

    #if extra_recipients != None and len(extra_recipients) > 0:
    #    log.debug(f"adding extra recipients { extra_recipients }")

    if args.send:
        #if extra_recipients != None and len(extra_recipients) > 0:
        #    message.bcc.add(extra_recipients)

        message.bcc.add(dr_config.to_test)
        message.bcc.add(dr_config.to_email)
        log.debug(f"sending { file_name } to { dr_config.to_email }")
        posting = f"<p>This message was sent to { dr_config.to_email }.  Please do *not* reply to the whole list</p>\n"
    else:
        log.debug(f"not sending { file_name } to { dr_config.to_email }")
        posting = \
f"""
<p>
DEBUG Version: not sent to the list
</p>
"""

    message.body = \
f"""
<!DOCTYPE html>
<html>
<meta http-equiv="Content-type" content="text/html" charset="UTF8" />
<title>DR{ dr_config.dr_id } { report_type }</title>
</head>
<body>

<h1>DR{ dr_config.dr_id } { report_type }</h1>
{ posting }

{ message_body }

</body>
</html>
"""


    message.subject = file_name
    message.attachments.add( file_name )

    try:
        message.send(save_to_sent_folder=True)
    except requests.RequestException as e:
        log.error(f"got an error: { e }, response json { e.response.json }")
        raise e


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
            description="tool to convert staffing reports into a more usefull form",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")
    parser.add_argument("--save", help="retain output file", action="store_true")
    parser.add_argument("--send", help="send emails out", action="store_true")
    parser.add_argument("--test-send", help="send emails out, but to the test email box", action="store_true")
    parser.add_argument("--dr-id", help="Identifier for the DRO; must match the staffing report", required=True, action="store")

    args = parser.parse_args()

    return args


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)
