"""
Happy path tests for VBA extraction
"""

import unittest, sys, os

from tests.test_utils import DATA_BASE_DIR
from os.path import join

if sys.version_info[0] <= 2:
    from oletools import olevba
else:
    from oletools import olevba3 as olevba

class TestVbaExtraction(unittest.TestCase):
    def test_vba_extraction(self):
        """ Test olevba.py with sample doc containing VBA """
        filename = join(DATA_BASE_DIR, 'olevba/harmless-with-vba.doc')
        vba_parser = olevba.VBA_Parser_CLI(filename, None, None, False)

        (_, _, vba_filename, vba_code) = vba_parser.extract_all_macros()[0]

        if isinstance(vba_code, bytes):
            vba_code = vba_code.decode('utf-8','backslashreplace')

        self.assertEqual(vba_filename, 'ThisDocument.cls')
        self.assertTrue('Hello World' in vba_code)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
