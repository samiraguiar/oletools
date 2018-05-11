""" Test some basic behaviour of msodde.py

Ensure that
- doc and docx are read without error
- garbage returns error return status
- dde-links are found where appropriate
"""

from __future__ import print_function

import unittest
from oletools import msodde
from tests.test_utils import DATA_BASE_DIR as BASE_DIR
from os.path import join
from traceback import print_exc


class TestReturnCode(unittest.TestCase):
    """ check return codes and exception behaviour (not text output) """
    def test_valid_doc(self):
        """ check that a valid doc file leads to 0 exit status """
        for filename in (
                'dde-test-from-office2003', 'dde-test-from-office2016',
                'harmless-clean', 'dde-test-from-office2013-utf_16le-korean'):
            self.do_test_validity(join(BASE_DIR, 'msodde',
                                       filename + '.doc'))

    def test_valid_docx(self):
        """ check that a valid docx file leads to 0 exit status """
        for filename in 'dde-test', 'harmless-clean':
            self.do_test_validity(join(BASE_DIR, 'msodde',
                                       filename + '.docx'))

    def test_valid_docm(self):
        """ check that a valid docm file leads to 0 exit status """
        for filename in 'dde-test', 'harmless-clean':
            self.do_test_validity(join(BASE_DIR, 'msodde',
                                       filename + '.docm'))

    def test_valid_xml(self):
        """ check that xml leads to 0 exit status """
        for filename in 'harmless-clean-2003.xml', 'dde-in-excel2003.xml', \
                'dde-in-word2003.xml', 'dde-in-word2007.xml':
            self.do_test_validity(join(BASE_DIR, 'msodde', filename))

    def test_invalid_none(self):
        """ check that no file argument leads to non-zero exit status """
        self.do_test_validity('', True)

    def test_invalid_empty(self):
        """ check that empty file argument leads to non-zero exit status """
        self.do_test_validity(join(BASE_DIR, 'basic/empty'), True)

    def test_invalid_text(self):
        """ check that text file argument leads to non-zero exit status """
        self.do_test_validity(join(BASE_DIR, 'basic/text'), True)

    def do_test_validity(self, args, expect_error=False):
        """ helper for test_valid_doc[x] """
        have_exception = False
        try:
            msodde.process_file(args, msodde.FIELD_FILTER_BLACKLIST)
        except Exception:
            have_exception = True
            print_exc()
        except SystemExit as exc:     # sys.exit() was called
            have_exception = True
            if exc.code is None:
                have_exception = False

        self.assertEqual(expect_error, have_exception,
                         msg='Args={0}, expect={1}, exc={2}'
                             .format(args, expect_error, have_exception))


class TestDdeLinks(unittest.TestCase):
    """ capture output of msodde and check dde-links are found correctly """

    @staticmethod
    def get_dde_from_output(output):
        """ helper to read dde links from captured output
        """
        return [o for o in output.splitlines()]

    def test_with_dde(self):
        """ check that dde links appear on stdout """
        filename = 'dde-test-from-office2003.doc'
        output = msodde.process_file(
            join(BASE_DIR, 'msodde', filename), msodde.FIELD_FILTER_BLACKLIST)
        self.assertNotEqual(len(self.get_dde_from_output(output)), 0,
                            msg='Found no dde links in output of ' + filename)

    def test_no_dde(self):
        """ check that no dde links appear on stdout """
        filename = 'harmless-clean.doc'
        output = msodde.process_file(
            join(BASE_DIR, 'msodde', filename), msodde.FIELD_FILTER_BLACKLIST)
        self.assertEqual(len(self.get_dde_from_output(output)), 0,
                         msg='Found dde links in output of ' + filename)

    def test_with_dde_utf16le(self):
        """ check that dde links appear on stdout """
        filename = 'dde-test-from-office2013-utf_16le-korean.doc'
        output = msodde.process_file(
            join(BASE_DIR, 'msodde', filename), msodde.FIELD_FILTER_BLACKLIST)
        self.assertNotEqual(len(self.get_dde_from_output(output)), 0,
                            msg='Found no dde links in output of ' + filename)

    def test_excel(self):
        """ check that dde links are found in excel 2007+ files """
        expect = ['DDE-Link cmd /c calc.exe', ]
        for extn in 'xlsx', 'xlsm', 'xlsb':
            output = msodde.process_file(
                join(BASE_DIR, 'msodde', 'dde-test.' + extn), msodde.FIELD_FILTER_BLACKLIST)

            self.assertEqual(expect, self.get_dde_from_output(output),
                             msg='unexpected output for dde-test.{0}: {1}'
                                 .format(extn, output))

    def test_xml(self):
        """ check that dde in xml from word / excel is found """
        for name_part in 'excel2003', 'word2003', 'word2007':
            filename = 'dde-in-' + name_part + '.xml'
            output = msodde.process_file(
                join(BASE_DIR, 'msodde', filename), msodde.FIELD_FILTER_BLACKLIST)
            links = self.get_dde_from_output(output)
            self.assertEqual(len(links), 1, 'found {0} dde-links in {1}'
                                            .format(len(links), filename))
            self.assertTrue('cmd' in links[0], 'no "cmd" in dde-link for {0}'
                                               .format(filename))
            self.assertTrue('calc' in links[0], 'no "calc" in dde-link for {0}'
                                                .format(filename))

    def test_clean_rtf_blacklist(self):
        """ find a lot of hyperlinks in rtf spec """
        filename = 'RTF-Spec-1.7.rtf'
        output = msodde.process_file(
            join(BASE_DIR, 'msodde', filename), msodde.FIELD_FILTER_BLACKLIST)
        self.assertEqual(len(self.get_dde_from_output(output)), 1413)

    def test_clean_rtf_ddeonly(self):
        """ find no dde links in rtf spec """
        filename = 'RTF-Spec-1.7.rtf'
        output = msodde.process_file(
            join(BASE_DIR, 'msodde', filename), msodde.FIELD_FILTER_DDE)
        self.assertEqual(len(self.get_dde_from_output(output)), 0,
                         msg='Found dde links in output of ' + filename)


if __name__ == '__main__':
    unittest.main()
