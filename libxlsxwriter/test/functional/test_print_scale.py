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

    def test_print_scale01(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_scale01')

    def test_print_scale02(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_print_scale02')
