/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2024, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "xlsxwriter.h"

int main() {

    lxw_workbook  *workbook  = workbook_new("test_repeat02.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    worksheet_set_paper(worksheet, 9);
    worksheet->vertical_dpi = 200;

    worksheet_repeat_columns(worksheet, 0, 0);

    worksheet_write_string(worksheet, CELL("A1"), "Foo" , NULL);

    return workbook_close(workbook);
}
