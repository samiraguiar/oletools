# encoding=utf8
""" Test the ability of msodde.py to detect different fields

Test against different files formats for documents and spreadsheets,
with multiple blacklisted fields
"""

import unittest
import re
from oletools import msodde
from tests.test_utils import OutputCapture, DATA_BASE_DIR as BASE_DIR
from os import listdir, path

EXPECTED_FIELDS_LEGACY = {
    re.compile(r'DDEAUTO'): ['cmd.exe', 'calc.exe'],
    # use word boundaries so DDE is not confused with DDEAUTO
    re.compile(r'\bDDE\b'): ['cmd.exe']
}

EXPECTED_FIELDS = dict(EXPECTED_FIELDS_LEGACY.items() + {
    re.compile(r'IMPORT'): ['cat.jpg'],
    re.compile(r'INCLUDETEXT'): ['doc.txt'],
    re.compile(r'\bINCLUDE\b'): ['doc.txt']
}.items())

# from https://stackoverflow.com/a/93029/2883579
CONTROL_CHARS = ''.join(map(unichr, range(0, 32) + range(127, 160)))
CONTROL_CHARS_REGEX = re.compile('[%s]' % re.escape(CONTROL_CHARS))

LEGACY_FILES_EXT = ['.doc', '.dot']


class TestWordDocuments(unittest.TestCase):
    test_data_dir = path.join(BASE_DIR, 'msodde', 'fields')

    def test_word(self):
        files_path = path.join(self.test_data_dir, 'document')
        files = listdir(files_path)

        for filename in files:
            file_extension = path.splitext(filename)[1]
            is_legacy_file = file_extension in LEGACY_FILES_EXT

            with OutputCapture() as capturer:
                msodde.main([path.join(files_path, filename), ])
                self.assertTrue(
                    self._all_fields_were_found(capturer, is_legacy_file))

    @staticmethod
    def _all_fields_were_found(capturer, legacy_file=False):
        # not all fields are shown for legacy Office files
        if legacy_file:
            fields_to_look = EXPECTED_FIELDS_LEGACY
        else:
            fields_to_look = EXPECTED_FIELDS

        dde_links_started = False
        for line in capturer:
            # remove non-printable characters (for Office 97-2003 files)
            line = CONTROL_CHARS_REGEX.sub('', line)

            if 'DDE Links:' in line:
                dde_links_started = True
                continue

            if not line.strip() or not dde_links_started:
                continue

            # remove from dict if the field was found and content matches
            for key in fields_to_look:
                if not key.search(line):
                    continue

                value = fields_to_look[key]

                if all(field_content in line for field_content in value):
                    fields_to_look.pop(key, None)
                    break

        # if everything was found `fields_to_look` should be empty
        return not fields_to_look


if __name__ == '__main__':
    unittest.main()
