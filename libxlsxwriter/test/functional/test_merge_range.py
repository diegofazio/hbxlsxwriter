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

    def test_merge_range01(self):
        self.run_exe_test('test_merge_range01')

    def test_merge_range02(self):
        self.run_exe_test('test_merge_range02')

    def test_merge_range03(self):
        self.run_exe_test('test_merge_range03')

    def test_merge_range04(self):
        self.run_exe_test('test_merge_range04')

    def test_merge_range05(self):
        self.run_exe_test('test_merge_range05')
