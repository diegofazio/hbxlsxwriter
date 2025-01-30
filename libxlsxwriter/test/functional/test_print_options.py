###############################################################################
#
# Tests for libxlsxwriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2014-2024, John McNamara, jmcnamara@cpan.org.
#

import base_test_class

class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with libxlsxwriter against a file created by Excel.

    """

    def test_print_options01(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_options01')

    def test_print_options02(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_options02')

    def test_print_options03(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_options03')

    def test_print_options04(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_options04')

    def test_print_options05(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_options05')

    def test_print_options06(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_options06')

    def test_print_options07(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_options07')
