
import datetime
import io
import logging
import pathlib

import O365.excel
import openpyxl


import neil_tools

#
# class to hold O365 related code to better modularize
#


class O365Config:
    """
    This class holds the code for processing the o365 side of the per-dro config file
    """

    def __init__(self, dr_config, account):

        self._dr_config = dr_config
        self._account = account
        # open the DR folder
        self._dr_site = account.sharepoint().get_site(dr_config.sharepoint_site, dr_config.dr_path)

        if self._dr_site is None:
            msg = f"Could not find sharepoint site '{ dr_config.sharepoint_site }' path '{ dr_config.dr_path }'"
            log.error(msg)
            raise Exception(msg)

        log.debug(f"site { self._dr_site } name '{ self._dr_site.display_name }'")

        self._dr_drive = self._dr_site.get_default_document_library()
        self._dr_folder = self.get_or_create_report_folder(dr_config.dr_folder_name)

        # open the template folder
        self._template_site = account.sharepoint().get_site(dr_config.sharepoint_site, dr_config.template_path)
        if self._template_site is None:
            msg = f"Could not find template site '{ dr_config.sharepoint_site }' path '{ dr_config.template_path }'"
            log.error(msg)
            raise Exception(msg)
        self._template_drive = self._template_site.get_default_document_library()
        self._template_folder = self._template_drive.get_item_by_path(dr_config.template_folder_name)

        # open the two files we care about, creating if necessary
        self.get_dr_file(dr_config.readme_file)
        self._report_config = self.get_dr_file(dr_config.report_config_file)
        log.debug(f"get_dr_file '{ dr_config.report_config_file }' returned { self._report_config }")


        # turn report_config into a spreadsheet
        config_stream = io.BytesIO()
        self._report_config.download(output=config_stream)
        self._config_wb = openpyxl.load_workbook(config_stream)



    #
    # try to open the report folder.  Create it if it doesn't exist
    #

    def get_or_create_report_folder(self, folder_name):
        try:
            # try to open the folder.  If it works: it exists
            folder = self.dr_drive.get_item_by_path(folder_name)
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:

                # the folder wasn't present.  Create it
                folder_path = pathlib.Path(folder_name)
                parent_path = folder_path.parent
                log.info(f"need to create per-dr folder '{ folder_name }' in parent '{ parent_path.as_posix() }'")
                if parent_path.as_posix() == '/':
                    parent = self.dr_drive.get_root_folder()
                else:
                    parent = self.dr_drive.get_item_by_path(parent_path.as_posix())

                log.debug(f"trying to create folder '{ folder_path.name }'")
                folder = parent.create_child_folder(folder_path.name)
            else:
                raise
        return folder


    #
    # get a file from the dr_folder.  If it doesn't exists: create it from the template folder
    # return an Entry for the file
    #

    def get_dr_file(self, file_name):

        file_path = pathlib.Path(self.dr_config.dr_folder_name).joinpath(file_name)
        template_path = pathlib.Path(self.dr_config.template_folder_name).joinpath(file_name)

        try:
            # try to open the folder.  If it works: it exists
            item = self.dr_folder.get_drive().get_item_by_path(file_path.as_posix())
        except requests.exceptions.HTTPError as e:
            if e.response.status_code != 404:
                raise

            # copy from the templates folder, which we assume exists
            template_item = self.template_folder.get_drive().get_item_by_path(template_path.as_posix())
            copy_op = template_item.copy(target=self.dr_folder, name=file_name)
            for status, percent_complete in coyp_op.check_status(delay=1):
                log.debug(f"copy_op: status { status } %complete { percent_complete}")
            item = copy_op.get_item()
            log.debug(f"copy_op: item { item }")


        return item


    #
    # save a copy of an openpyxl workbook to sharepoint/onedrive
    #
    def save_report_file(self, wb, file_path):

        dr_folder = self.dr_folder

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
    # initialize the o365.excel workbook coresponding to the Report Config spreadsheet
    #

    def init_config_wb(self, roster_file_name, roster_sps_file_name, run_date, report_date):

        # get o365 DriveItem for the config workbook
        report_config = self.report_config
        dr_config = self.dr_config


        # get the config workbook
        self._config_wb = O365.excel.WorkBook(report_config, use_session=False)

        # start with the report status worksheet
        self._ws_status = self._config_wb.get_worksheet(dr_config.sheet_report_status)
        log.debug(f"ws_status { self._ws_status }")
                                                                                                                                   # insert a blank row so newest is on top
        old_status_range = self._ws_status.get_range("A2:E2")
        old_status_values = old_status_range.values
        new_status_range = old_status_range.insert_range("down")

        # and add the data
        new_status_range.values = [[ run_date.isoformat(), report_date.isoformat(),
                                    "Started", roster_file_name, roster_sps_file_name ]]
        new_status_range.update()

        # return the time of the last staffing report
        last_report_date_string = old_status_values[0][1]
        if isinstance(last_report_date_string, str) and last_report_date_string != "":
            last_report_date = datetime.datetime.fromisoformat(last_report_date_string)
        else:
            last_report_date = None
        log.debug(f"last report time: { last_report_date } str '{ last_report_date_string }'")
        return last_report_date

    #
    # update the report status field
    #

    def update_report_status(self, new_status):
        status_range = self._ws_status.get_range("C2:C2")
        status_range.values = [[ new_status ]]
        status_range.update()
        log.debug(f"update_report_status: new status '{ new_status }'")



    #
    # initialize and clear the recipient worksheet
    #

    def init_recipient_sheet(self):
        self._ws_recip = self._config_wb.get_worksheet(self.dr_config.sheet_current_recipients)
        log.debug(f"ws_recip { self._ws_recip }")

        recips_used_range = self._ws_recip.get_used_range()
        recips_used_range.clear()

        # build the array to update result
        values = [ [ "Report Type", "Name", "Email", "Gap" ] ]      # title row

        update_range = f"A1:D{len(values)}"
        log.debug(f"current_recipients update range is '{ update_range }'")
        recip_update_range = self._ws_recip.get_range(update_range)
        recip_update_range.values = values
        recip_update_range.update()


    #
    # add a batch of recipients to the recipient sheet
    #

    def update_recipient_sheet(self, mailing_list, report_type):

        ws = self._ws_recip

        # mailing list entries are either a str (email) or a dict with roster row data
        values = []
        for entry in mailing_list:
            if type(entry) == str:
                values.append( [ report_type, None, entry, None ] )
            else:
                values.append( [ report_type, entry['Name'], entry['Email'], entry['GAP(s)'] ] )

        recips_used_range = ws.get_used_range()
        row_start = recips_used_range.row_index + recips_used_range.row_count
        col_index = recips_used_range.column_index
        col_count = recips_used_range.column_count
        log.debug(f"update_recipient_sheet: { report_type } row_start { row_start } col_index { col_index } col_count { col_count }")

        update_range = f"A{ row_start + 1 }:D{ row_start + len(values)}"
        log.debug(f"current_recipients update range is '{ update_range }', len { len(values) }")
        recip_update_range = ws.get_range(update_range)
        recip_update_range.values = values
        recip_update_range.update()



    @property
    def account(self):
        return self._account

    @property
    def dr_config(self):
        return self._dr_config

    @property
    def dr_site(self):
        return self._dr_site

    @property
    def dr_folder(self):
        return self._dr_folder

    @property
    def dr_drive(self):
        return self._dr_drive

    @property
    def config_wb(self):
        return self._config_wb

    @property
    def report_config(self):
        return self._report_config


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)

