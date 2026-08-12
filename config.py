# config.py

TOKEN_FILENAME = 'o365_token.txt'

PROGRAM_EMAIL = 'DR-Report-Automation@redcross.org'
EMAIL_TO_TEST = 'generic@askneil.com'

SHAREPOINT_SITE = 'americanredcross.sharepoint.com'
TEMPLATE_PATH = '/teams/dr-staffing-reports'
TEMPLATE_FOLDER = '/templates'
DR_FOLDER = '/Report Automation'
REPORTS_FOLDER = "Reports"

README_FILE = '000_README_ME_FIRST.md'
REPORT_CONFIG_FILE = 'Report Config.xlsx'

# sheet names
SHEET_PER_GAP = 'Per GAP config'
SHEET_EXTRA_RECIPIENTS = 'Extra Recipients'
SHEET_CURRENT_RECIPIENTS = 'Current Recipients'
SHEET_REPORT_STATUS = 'Report Status'

# named tables
TABLE_PER_GAP = 'PerGap'
TABLE_EXTRA_RECIPIENTS = 'ExtraRecipients'

# default value for daily checkin nags
LATE_CHECKIN_THRESHOLD = 3


_DR_CONFIGURATIONS = {}

class DRConfig:
    def __init__(self, dr_id, to_email, dr_path, subject_match_string=None, dr_folder=DR_FOLDER, from_email=None):
        self._dr_id = dr_id
        self._from_email = from_email
        self._to_email = to_email
        self._to_test = EMAIL_TO_TEST
        self._dr_path = dr_path
        self._dr_folder = dr_folder
        self._sharepoint_site = SHAREPOINT_SITE
        self._template_path = TEMPLATE_PATH
        self._template_folder = TEMPLATE_FOLDER
        self._reports_folder = REPORTS_FOLDER

        self._readme_file = README_FILE
        self._report_config_file = REPORT_CONFIG_FILE

        self._sheet_per_gap = SHEET_PER_GAP
        self._sheet_extra_recipients = SHEET_EXTRA_RECIPIENTS
        self._sheet_current_recipients = SHEET_CURRENT_RECIPIENTS
        self._sheet_report_status = SHEET_REPORT_STATUS

        self._table_per_gap = TABLE_PER_GAP
        self._table_extra_recipients = TABLE_EXTRA_RECIPIENTS

        self._subject_match_string = subject_match_string

        self._late_checkin_threshold = LATE_CHECKIN_THRESHOLD

        _DR_CONFIGURATIONS[self.dr_id] = self

    @property
    def dr_id(self):
        return self._dr_id

    @property
    def from_email(self):
        return self._from_email

    @property
    def to_email(self):
        return self._to_email

    @property
    def to_test(self):
        return self._to_test

    @property
    def dr_path(self):
        return self._dr_path

    @property
    def dr_folder(self):
        return self._dr_folder

    @property
    def reports_folder(self):
        return self._reports_folder

    @property
    def sharepoint_site(self):
        return self._sharepoint_site

    @property
    def template_path(self):
        return self._template_path

    @property
    def template_folder(self):
        return self._template_folder

    @property
    def readme_file(self):
        return self._readme_file

    @property
    def report_config_file(self):
        return self._report_config_file

    @property
    def sheet_extra_recipients(self):
        return self._sheet_extra_recipients

    @property
    def sheet_per_gap(self):
        return self._sheet_per_gap

    @property
    def sheet_current_recipients(self):
        return self._sheet_current_recipients

    @property
    def sheet_report_status(self):
        return self._sheet_report_status

    @property
    def table_per_gap(self):
        return self._table_per_gap

    @property
    def table_extra_recipients(self):
        return self._table_extra_recipients

    @property
    def subject_match_string(self):
        return self._subject_match_string

    @property
    def late_checkin_threshold(self):
        return self._late_checkin_threshold

    @staticmethod
    def lookup_dr(dr_id):
        if dr_id not in _DR_CONFIGURATIONS:
            return None
        return _DR_CONFIGURATIONS[dr_id]


# now create the DROs
DRConfig('458-2026', 'neil.katin@redcross.org', '/teams/DR458-26NWRegionWildfires-Workforce', subject_match_string='DR458-2026Automated Workforce Reports')

