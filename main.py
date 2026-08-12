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
import pathlib
import requests
import requests.exceptions
import textwrap

import xlrd
import dotenv

import openpyxl
import openpyxl.utils
import openpyxl.styles
import openpyxl.styles.colors
import openpyxl.writer.excel
import openpyxl.utils.cell
import O365.excel

import config as config_static
import neil_tools
import arc_o365
from neil_tools import spreadsheet_tools


# index of column label row, origin zero
STAFF_ROSTER_LABEL_ROW = 5
ORIG_SHEET_NAME = "Orig"
REPORT_DATE_FORMAT = "%Y-%m-%d %H-%M-%S %Z"


NOW = datetime.datetime.now().astimezone()
NOW_NO_TZ = datetime.datetime.now()


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

    report_dict = o365.fetch_workforce_reports(dr_config.dr_id, subject_match_string=dr_config.subject_match_string)
    report_date = report_dict['created']
    report_date_stamp = report_date.strftime(REPORT_DATE_FORMAT)
    log.debug(f"report date is '{ report_date }', stamp '{ report_date_stamp }'")

    errors = False
    book_out = openpyxl.Workbook()

    # do the 'orig' roster first so it is at the end of the list
    sheet_orig = read_roster(book_out, ORIG_SHEET_NAME, report_dict['Staff Roster - Cumulative'], STAFF_ROSTER_LABEL_ROW, ROSTER_FIXUPS)
    sheet_roster1 = copy_sheet(dr_config, book_out, sheet_orig, STAFF_ROSTER_LABEL_ROW, "Roster1", filter_row_active, ROSTER_FIXUPS)

    # delete column 'I': the Released column
    # first delete the column
    sheet_roster1.delete_cols(openpyxl.utils.cell.column_index_from_string('I'), 1)
    # adjust the table size
    sheet_name, table_range = find_table_by_name(book_out, 'Roster1')
    log.debug(f"sheet { sheet_name } range { table_range }")

    # get the range in terms of min/max
    # tuple is min_col, min_row, max_col, max_row
    tuple_range = openpyxl.utils.cell.range_boundaries(table_range)
    log.debug(f"table range_boundaries { tuple_range }")
    # convert back to A1:Z999 style format, reducing the max column by one for the deleted column
    min_col = openpyxl.utils.get_column_letter(tuple_range[0])
    min_row = tuple_range[1]
    max_col = openpyxl.utils.get_column_letter(tuple_range[2] -1)
    max_row = tuple_range[3]
    table_ref = f"{min_col}{min_row}:{max_col}{max_row}"
    log.debug(f"resizing table: table_ref '{ table_ref }'")
    table = openpyxl.worksheet.table.Table(displayName='Roster1', ref=table_ref)
    # delete the old table
    del sheet_roster1.tables['Roster1']
    # add the new one
    sheet_roster1.add_table(table)

    # and copy it to a new sheet to fix up all the column header widths
    sheet_roster = copy_sheet(dr_config, book_out, sheet_roster1, 0, "Roster", {}, ROSTER_FIXUPS)
    del book_out['Roster1']


    read_roster(book_out, 'StaffRequests', report_dict['Open Staff Requests'], 1, ROSTER_FIXUPS)
    read_roster(book_out, 'Shifts', report_dict['DRO Shift Tool - Shift Registrant Details'], 3, SHIFTS_FIXUPS)
    read_roster(book_out, 'Air', report_dict['Air Travel Roster'], 2, AIR_FIXUPS, freeze_col="C", suppress_columns={'V':True})
    read_roster(book_out, 'Arrival', report_dict['Arrival Roster'], 5, ARRIVAL_FIXUPS, suppress_columns={'Z':True})

    sps_sheets = {
        "Late_Checkin": filter_row_checkin,
        "Need_SMS": filter_row_sms,
        "Needs_Sup": filter_row_needs_sup,
        "Days_2": filter_row_2_days,
        "Outprocess": filter_row_overstayed,
    }

    all_sheets = {
        "OM": filter_row_om,
        "WF": filter_row_wf,
        "IP": filter_row_ip,
        "ER": filter_row_er,
        "LOG": filter_row_log,
        "CC": filter_row_cc,
        "MC": filter_row_mc,
        }

    # make all the SPS sheets
    for sheet_name, filter in sps_sheets.items():
        copy_sheet(dr_config, book_out, sheet_roster, 0, sheet_name, filter, ROSTER_FIXUPS, sheet_color='99ff99')
    for sheet_name, filter in all_sheets.items():
        copy_sheet(dr_config, book_out, sheet_roster, 0, sheet_name, filter, ROSTER_FIXUPS)

    # remove the default sheet in a new wb that we don't need
    del book_out['Sheet']

    # move reoster to beginning of workbook
    # we should be using workbook.move_sheet(), but that didn't seem to work
    book_out.remove(sheet_roster)
    book_out._add_sheet(sheet_roster, index=0)

    # name of saved file
    roster_file_name = f"DR{ dr_config.dr_id } Staffing Report { report_date_stamp }.xlsx"
    roster_sps_file_name = f"DR{ dr_config.dr_id } SPS Report { report_date_stamp }.xlsx"

    # now figure out who we should send to
    config_tables = open_report_automation(dr_config, o365.account)
    mailing_list = run_gap_patterns(dr_config, config_tables, book_out, 'Roster')

    update_config_wb(dr_config, config_tables, mailing_list, roster_file_name, roster_sps_file_name, report_date_stamp)

    if errors:
        sys.exit(1)

    if args.send or args.test_send or args.save:
        log.debug(f"saving to { roster_file_name }")
        book_out.save(roster_file_name)
        save_report_file(dr_config, config_tables, book_out, dr_config.reports_folder + "/" + roster_file_name)

    if args.send or args.test_send:
        send_roster(dr_config, args, o365.account, roster_file_name, "SPS Report", report_date)

    if not args.save and (args.send or args.test_send):
        log.debug(f"removing { roster_file_name }")
        os.remove(roster_file_name)


#
# temporary code to test finding a drive id by name
#
#
def open_report_automation(dr_config, account):

    # open the DR folder
    dr_site = account.sharepoint().get_site(dr_config.sharepoint_site, dr_config.dr_path)

    if dr_site is None:
        msg = f"Could not find sharepoint site '{ dr_config.sharepoint_site }' path '{ dr_config.dr_path }'"
        log.error(msg)
        raise Exception(msg)

    log.debug(f"site { dr_site } name '{ dr_site.display_name }'")

    dr_drive = dr_site.get_default_document_library()
    dr_folder = get_or_create_report_folder(dr_config, dr_drive, dr_config.dr_folder)
    #reports_folder = get_or_create_report_folder(dr_config, dr_drive, dr_config.reports_folder)

    # open the template folder
    template_site = account.sharepoint().get_site(dr_config.sharepoint_site, dr_config.template_path)
    if template_site is None:
        msg = f"Could not find template site '{ dr_config.sharepoint_site }' path '{ dr_config.template_path }'"
        log.error(msg)
        raise Exception(msg)
    template_drive = template_site.get_default_document_library()
    template_folder = template_drive.get_item_by_path(dr_config.template_folder)

    # open the two files we care about, creating if necessary
    get_dr_file(dr_config, dr_config.readme_file, template_folder, dr_folder)
    report_config = get_dr_file(dr_config, dr_config.report_config_file, template_folder, dr_folder)
    log.debug(f"get_dr_file '{ dr_config.report_config_file }' returned { report_config }")


    # turn report_config into a spreadsheet
    config_stream = io.BytesIO()
    report_config.download(output=config_stream)
    config_wb = openpyxl.load_workbook(config_stream)

    config_tables = { 'config_wb': config_wb,
                     'dr_folder': dr_folder,
                     'reports_folder_name': dr_config.reports_folder,
                     'report_config': report_config }
    # get the by-gap table
    for table_name in [ dr_config.table_per_gap, dr_config.table_extra_recipients ]:
        config_tables[table_name] = read_table_to_dict(config_wb, table_name)

    return config_tables




#
# try to open the report folder.  Create it if it doesn't exist
#

def get_or_create_report_folder(dr_config, dr_drive, folder_name):
    try:
        # try to open the folder.  If it works: it exists
        folder = dr_drive.get_item_by_path(folder_name)
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 404:

            # the folder wasn't present.  Create it
            folder_path = pathlib.Path(folder_name)
            parent_path = folder_path.parent
            log.info(f"need to create per-dr folder '{ folder_name }' in parent '{ parent_path.as_posix() }'")
            if parent_path.as_posix() == '/':
                parent = dr_drive.get_root_folder()
            else:
                parent = dr_drive.get_item_by_path(parent_path.as_posix())

            log.debug(f"trying to create folder '{ folder_path.name }'")
            folder = parent.create_child_folder(folder_path.name)
        else:
            raise
    return folder


#
# get a file from the dr_folder.  If it doesn't exists: create it from the template folder
# return an Entry for the file
#

def get_dr_file(dr_config, file_name, template_folder, dr_folder):

    file_path = pathlib.Path(dr_config.dr_folder).joinpath(file_name)
    template_path = pathlib.Path(dr_config.template_folder).joinpath(file_name)

    try:
        # try to open the folder.  If it works: it exists
        item = dr_folder.get_drive().get_item_by_path(file_path.as_posix())
    except requests.exceptions.HTTPError as e:
        if e.response.status_code != 404:
            raise

        # copy from the templates folder, which we assume exists
        template_item = template_folder.get_drive().get_item_by_path(template_path.as_posix())
        copy_op = template_item.copy(target=dr_folder, name=file_name)
        for status, percent_complete in coyp_op.check_status(delay=1):
            log.debug(f"copy_op: status { status } %complete { percent_complete}")
        item = copy_op.get_item()
        log.debug(f"copy_op: item { item }")


    return item


#
# run through all the sheets, looking for the specified table.
#
# return the sheet name and the table range
#

def find_table_by_name(wb, name):
    for ws in wb.worksheets:
        for table_name, table_range in ws.tables.items():
            if name == table_name:
                return ws.title, table_range
    return None, None


#
# read a table into a dict, with the keys being the table headers
#

def read_table_to_dict(wb, table_name):

    # find the table
    sheet_name, table_range = find_table_by_name(wb, table_name)
    #log.debug(f"table { table_name } sheet { sheet_name } range { table_range }")

    # get the range in terms of min/max
    # tuple is min_col, min_row, max_col, max_row
    tuple_range = openpyxl.utils.cell.range_boundaries(table_range)
    #log.debug(f"table range_boundaries { tuple_range }")

    ws = wb[sheet_name]

    row_list = []
    for row in ws.iter_rows(
            min_col=tuple_range[0], min_row=tuple_range[1],
            max_col=tuple_range[2], max_row=tuple_range[3], values_only=True):
        
        if not all_none(row):
            row_list.append(row)

    #log.debug(f"row_list: { row_list }")

    table_dict = spreadsheet_tools.matrix_to_object_array(row_list)
    #log.debug(f"table_dict: { table_dict }")

    return table_dict

#
# returns True if all elements in a list are None
#

def all_none(row):
    return all(x is None for x in row)


#
# go through the config spreadsheet and figure out who this report should be sent to
#

def run_gap_patterns(dr_config, config_tables, wb, roster_table_name):

    roster_table_dicts = read_table_to_dict(wb, roster_table_name)
    config_wb = config_tables['config_wb']
    per_gap_dicts = config_tables[dr_config.table_per_gap]
    extra_recipients = config_tables[dr_config.table_extra_recipients]

    # start by including everyone on the Extra Recipients list who is "include"
    people_to_include = list( filter(lambda x: x is not None,
            map(lambda x: x['Email'] if x['Type'] == "include" else None,
            extra_recipients)) )

    log.debug(f"people_to_include part 1: { people_to_include }")

    per_gap_dicts_to_re(per_gap_dicts)

    for person_dict in roster_table_dicts:
        #log.debug(f"roster_table_dicts: { person_dict }")
        email = person_dict['Email']

        # if email is in extra_recipients (no matter if include or exclude) then
        # skip further processing
        #
        # if include: they are already there
        # if exclude: prevent them from being added
        if any(filter(lambda x: x['Email'] == email, extra_recipients)):
            log.debug(f"skipping email { email } because it is in extra_recipients")
            continue

        match_per_gap(person_dict, per_gap_dicts, people_to_include)

    log.debug(f"after gap filtering: including { len(people_to_include) } people")
    return people_to_include

#
# see if we should include this email based on the gap of the person
#
def match_per_gap(person_dict, gap_dicts, people_to_include):

    email = person_dict['Email']
    gap = person_dict['GAP(s)']

    for e in gap_dicts:
        gap_re = e['re']

        if gap_re.fullmatch(gap) is not None:
            if e['Type'] == 'include':
                #log.debug(f"Including { email } gap { gap } based on { e['GAP'] }")
                people_to_include.append(person_dict)
            else:
                #log.debug(f"Excluding { email } gap { gap } based on { e['GAP'] }")
                pass

            # include or exclude: we're done with matching
            return


    log.debug(f"no match for { email } gap { gap }") 



#
# go through all the per-gap-dict entries and turn the glob expressions into
# compiled regular expressions
#

def per_gap_dicts_to_re(gap_dicts):

    for e in gap_dicts:
        #log.debug(f"per_gap_dicts_to_re: e { e }")
        gap = e['GAP']

        # we have something that looks like 'om-*' or 'mc-sh-mn'
        # steps to do:
        #    upper case everything
        #    turn - to /
        #    turn * to .*
        #    anchor at end to match whole string
        orig_gap = gap
        gap = gap.upper()
        gap = re.sub(r'\*', '.*', gap)
        gap = re.sub(r'-', '/', gap)

        #log.debug(f"orig_gap '{ orig_gap }' gap '{ gap }'")

        e['re'] = re.compile(gap, re.DOTALL)


#
# save results to the config workbook
#
def update_config_wb(dr_config, config_tables, mailing_list, roster_file_name, roster_sps_file_name, report_date):

    # get o365 DriveItem for the config workbook
    report_config = config_tables['report_config']


    # get the config workbook
    o365_wb = O365.excel.WorkBook(report_config, use_session=False)

    # start with the report status worksheet
    ws_status = o365_wb.get_worksheet(dr_config.sheet_report_status)
    log.debug(f"ws_status { ws_status }")

    # insert a blank row so newest is on top
    old_status_range = ws_status.get_range("A2:E2")
    new_status_range = old_status_range.insert_range("down")
    
    # and add the data
    new_status_range.values = [[ NOW.strftime(REPORT_DATE_FORMAT), report_date,
                                "Success", roster_file_name, roster_sps_file_name ]]
    new_status_range.update()

    # now do the Current Recipients worksheet
    ws_recip = o365_wb.get_worksheet(dr_config.sheet_current_recipients)
    log.debug(f"ws_recip { ws_recip }")

    recips_used_range = ws_recip.get_used_range()
    recips_used_range.clear()

    # build the array to update result
    values = [ [ "Name", "Email", "Gap" ] ]      # title row

    # mailing list entries are either a str (email) or a dict with roster row data
    for entry in mailing_list:
        if type(entry) == str:
            values.append( [ None, entry, None ] )
        else:
            values.append( [ entry['Name'], entry['Email'], entry['GAP(s)'] ] )

    update_range = f"A1:C{len(values)}"
    log.debug(f"current_recipients update range is '{ update_range }'")
    recip_update_range = ws_recip.get_range(update_range)
    recip_update_range.values = values
    recip_update_range.update()





#
# not working since you can't write to a file that is open somewhere
#
def save_report_file(dr_config, config_tables, wb, file_path):

    dr_folder = config_tables['dr_folder']

    # step three: write the file back to the dr_folder in sharepoint
    stream = io.BytesIO()
    wb.save(stream)
    content = stream.getvalue()
    content_len = len(content)
    stream = io.BytesIO(content)

    # actually do the upload
    retval = dr_folder.upload_file(dr_folder, item_name=file_path, stream=stream, stream_size=content_len)
    log.debug(f"save_report_file: file { file_path } retval { retval }")



#
# simple function that returns True if all elements of a list are None
#

def dt_convert(c):

    # definitely not a date
    if c == '' or c is None:
        return ''

    # already converted
    if isinstance(c, datetime.datetime):
        return c

    #new = spreadsheet_tools.excel_to_dt(c)
    new = xlrd.xldate_as_datetime(c, 0)

    #log.debug(f"dt_convert: orig '{ c }' new '{ new.strftime(REPORT_DATE_FORMAT) }'")
    return new


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
        'Assigned': { 'convert_value': dt_convert, 'number_format': "yyyy-mm-dd", },
        'Checked in': { 'convert_value': dt_convert, 'number_format': "yyyy-mm-dd", },
        'Released': { 'convert_value': dt_convert, 'number_format': "yyyy-mm-dd", },
        'Travel home': { 'convert_value': dt_convert, 'number_format': "yyyy-mm-dd", },
        'Last Daily Checkin': { 'convert_value': dt_convert, 'number_format': "yyyy-mm-dd", },
        'DaysRemain': { 'width': 5, 'number_format': "##0",
                       'alignment': RIGHT_ALIGNED,
                       #'convert_value': lambda x: x if x is not None and isinstance(x, int)) else x if  x == '' or x == 'n/a' else int(x),
                       'convert_value': lambda x: None if x is None else x if isinstance(x, int) else x if  x == '' or x == 'n/a' else int(x),
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
        'Arrive date': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd",
                     },
        'Flight Arrival Date/Time': { 'convert_value': dt_convert,
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

        'Last action date': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'Exp Arrival': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd",
                     },
        'Departure time': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'Arrival time': { 'convert_value': dt_convert,
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

        'Start Date': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd",
                     },
        'Start Time': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'End Date': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'Date Registered/Last Changed': { 'convert_value': dt_convert,
                     'number_format': "yyyy-mm-dd hh:mm",
                     'width': 18,
                     },
        'End Time': { 'convert_value': dt_convert,
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

# see if the responder is missing daily checkins.
# the rule: last checkin over 3 days, or if never checked in: DRO checkin date over 3 days
def filter_row_checkin_func(val, row, row_name_map, dr_config):

    checkin_warn_threshold = dr_config.late_checkin_threshold
    dt_threshold = NOW_NO_TZ - datetime.timedelta(days=checkin_warn_threshold)
    if isinstance(val, datetime.datetime):
        if val < dt_threshold:
            return True
    else:
        c = row_name_map['Checked in']
        val = row[c]
        if isinstance(val, datetime.datetime):
            if val < dt_threshold:
                return True

    return False


filter_row_active = { 'Released': lambda x: x == '' }
filter_row_overstayed = { 'DaysRemain': lambda x: x != '' and x != 'n/a' and x is not None and int(x) <= 0 }
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

filter_row_checkin = { 'Last Daily Checkin': filter_row_checkin_func, 'extra_args': True }


def filter_row(row, row_name_map, filter_defs, dr_config):
    # check all the conditions in the filter; if they all pass include the row

    for name, func in filter_defs.items():
        if name == 'extra_args':
            # pseudo argument, not a column name.  Ignore
            continue

        #log.debug(f"filter_row: name '{ name }'")

        # ignore filtered columns
        if name in row_name_map:
            c = row_name_map[name]
            val = row[c]
            #log.debug(f"filter_row: name { name } c { c } val '{ val }' ")
            if 'extra_args' in filter_defs:
                pass
                if not func(val, row, row_name_map, dr_config):
                    #log.debug(f"filter_row: name { name } val '{ val }' returning False")
                    return False
            else:
                if not func(val):
                    #log.debug(f"filter_row: name { name } val '{ val }' returning False")
                    return False
        
    #log.debug(f"filter_row: returning True")
    return True


def read_roster(book_out, sheet_name, file_contents: str, label_row: int, fixups: dict, freeze_col: str = "B", suppress_columns: dict[str] = {}) -> openpyxl.worksheet.worksheet:
    
    book_in = xlrd.open_workbook(file_contents=file_contents)
    sheet_in = book_in.sheet_by_index(0)

    #log.debug(f"sheet name { sheet_in.name } rows { sheet_in.nrows } cols { sheet_in.ncols }")

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
def copy_sheet(dr_config, wb, sheet_orig, label_row, sheet_name, filters, fixups,
               sheet_color: str = None,
               suppress_columns: dict[str] = {}):
    
    #log.debug(f"copy_sheet: sheet_name { sheet_name } label_row { label_row }")
    #sheet_new = wb.create_sheet(sheet_name, len(wb.sheetnames)-1)
    sheet_new = wb.create_sheet(sheet_name, 0)
    if sheet_color is not None:
        sheet_new.sheet_properties.tabColor = sheet_color

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
            include_row = filter_row(row_values, column_name_map, filters, dr_config)
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
def send_roster(dr_config, args, account, file_name, report_type, report_date):

    warn_days = 2
    if report_date < NOW - datetime.timedelta(days=warn_days):
        date_warning = textwrap.dedent(
                f"""
                <p>
                <span style='background-color:yellow;'>WARNING: staff report is more than { warn_days } days old<span>.
                </p>
                """)
    else:
        date_warning = ""

    if report_type == "SPS Report":
        pass
        sps_section = textwrap.dedent(
                f"""
                <p>
                This is the SPS version of this report
                which several additional worksheets to slightly ease the SPS workflow.
                </p>

                <div style="padding-left:20px;">

                    <p>
                    The Outprocess worksheet has people with zero days (or negative days) of remaining time on the DR.
                    They should be contacted to outprocess or extend their deployment
                    </p>

                    <p>
                    The Days_2 worksheet has people with 2 days left on their deployment.
                    It is common to reach out to these folks with a template with outprocessing instructions
                    </p>

                    <p>
                    The Needs_Sup sheet shows people with their supervisor set to Needs Supervisor
                    </p>

                    <p>
                    The Need_SMS sheet shows people who have not opted in to DCS Text Messaging;
                    they should be contacted to opt in so they can receive DRIS and other text messages.
                    </p>

                    <p>
                    Late_checkin show people who haven't done a daily checkin in more than
                    { dr_config.late_checkin_threshold } days
                    </p>
                </div>

                """)
    else:
        sps_section = ""

    message_body = textwrap.dedent(
        f"""
        <p>
        This report is based on the staff roster dated { report_date.strftime(REPORT_DATE_FORMAT) }.
        </p>

        { date_warning }

        <p>
        Hello everyone.  This is a spreadsheet based on the automated staffing reports, but reorganized to hopefully be easier to use.
        </p>


        <p>
        It has the same data as the normal automated report, but has been reformatted for easier use.
        All the sheets have been made with tables, so the columns are sortable and filterable.
        The sheets have been 'frozen', so column headers and names are always visible,
        even when scrolling in the worksheet.
        </p>

        <p>
        There is a worksheet called "Roster" which has all responders who have not been out-processed appear.
        By default this is sorted by GAP, but you can easily re-sort it using the column headers
        </p>

        <p>
        In addition there is a worksheet for every Group, with the responders in that group in that worksheet.
        </p>

        <p>
        Finally there is a sheet for all the other automated reports,
        as well as the original version of the staff roster (called "Orig") that has outprocessed people also.
        </p>

        { sps_section }

        <p>
        DR { dr_config.dr_id } Staff Planning and Support
        </p>

        <p>
        If you have any questions or suggestions about the report or its contents, feel free to contact
        <a href='mailto:dr-report-automation@redcross.org'>dr-report-automation@redcross.org</a>.
        </p>

        """)

    send_report_common(dr_config, args, account, file_name, report_type, message_body)




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
