# config.py

TOKEN_FILENAME = 'o365_token.txt'

PROGRAM_EMAIL = 'DR-Report-Automation@redcross.org'
EMAIL_TO_TEST = 'generic@askneil.com'

_DR_CONFIGURATIONS = {}

class DRConfig:
    def __init__(self, dr_id, to_email, from_email=None):
        self._dr_id = dr_id
        self._from_email = from_email
        self._to_email = to_email
        self._to_test = EMAIL_TO_TEST

        _DR_CONFIGURATIONS[self.dr_id] = self

    @property
    def dr_id(self):
        return self._dr_id

    @property
    def to_email(self):
        return self._to_email

    @property
    def to_test(self):
        return self._to_test

    @property
    def from_email(self):
        return self.from_email

    @staticmethod
    def lookup_dr(dr_id):
        if dr_id not in _DR_CONFIGURATIONS:
            return None
        return _DR_CONFIGURATIONS[dr_id]


# now create the DROs
DRConfig("458-2026", "neil.katin@redcross.org")

