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

    def test_chart_table01(self):
        self.run_exe_test('test_chart_table01')

    def test_chart_table02(self):
        self.run_exe_test('test_chart_table02')

    def test_chart_table03(self):
        self.ignore_elements = {'xl/charts/chart1.xml': ['<a:pPr']}
        self.run_exe_test('test_chart_table03')
