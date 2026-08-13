
import io
import logging
import pathlib

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

