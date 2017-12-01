# encoding=utf8
""" Test the ability of msodde.py to detect different fields

Test against different files formats for documents and spreadsheets,
with multiple blacklisted fields
"""

import unittest
from oletools import msodde
from tests.test_utils import OutputCapture, DATA_BASE_DIR as BASE_DIR
from os import listdir, path

EXPECTED_FIELDS = [
    'DDEAUTO c:\\\\windows\\\\system32\\\\cmd.exe "/k calc.exe \'\'""\'"\'"\'"\'"\'"',
    u'DDE c:\\\\windows\\\\system32\\\\cmd.exe "ЁЂ🙇التَّطْبِي؀؁؂؃؄؅؜۝܏᠎''',
    'IMPORT  "C:\\\\Users\\\\foo\\\\Desktop\\\\cat.jpg" \* MERGEFORMAT',
    'INCLUDEPICTURE  "C:\\\\Users\\\\foo\\\\Desktop\\\\cat.jpg" \* MERGEFORMATINET',
    'INCLUDETEXT  "C:\\\\Users\\\\foo\\\\Desktop\\\\doc.txt" \* MERGEFORMAT',
    'INCLUDE  "C:\\\\Users\\\\foo\\\\Desktop\\\\doc.txt" \* MERGEFORMAT'
]

class TestWordDocuments(unittest.TestCase):
    test_data_dir = path.join(BASE_DIR, 'msodde', 'fields')

    def test_word(self):
        files_path = path.join(self.test_data_dir, 'document')
        files = listdir(files_path)

        for filename in files:
            with OutputCapture() as capturer:
                msodde.main([path.join(files_path, filename), ])
                self.assertTrue(self.all_fields_were_found(capturer))

    def all_fields_were_found(self, capturer):
        fields_to_look = EXPECTED_FIELDS

        for line in capturer:
            if not line.strip():
                continue

            fields_to_look = [f for f in fields_to_look if f not in line]
        
        return not fields_to_look


if __name__ == '__main__':
    unittest.main()
