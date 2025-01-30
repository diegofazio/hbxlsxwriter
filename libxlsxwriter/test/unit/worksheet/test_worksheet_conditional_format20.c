/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2024, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/xlsxwriter/worksheet.h"
#include "../../../include/xlsxwriter/shared_strings.h"

// Test assembling a complete Worksheet file.
CTEST(worksheet, worksheet_condtional_format20) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
              "<dimension ref=\"A1:A4\"/>"
              "<sheetViews>"
                "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
              "</sheetViews>"
              "<sheetFormatPr defaultRowHeight=\"15\"/>"
              "<sheetData>"
                "<row r=\"1\" spans=\"1:1\">"
                  "<c r=\"A1\">"
                    "<v>10</v>"
                  "</c>"
                "</row>"
                "<row r=\"2\" spans=\"1:1\">"
                  "<c r=\"A2\">"
                    "<v>20</v>"
                  "</c>"
                "</row>"
                "<row r=\"3\" spans=\"1:1\">"
                  "<c r=\"A3\">"
                    "<v>30</v>"
                  "</c>"
                "</row>"
                "<row r=\"4\" spans=\"1:1\">"
                  "<c r=\"A4\">"
                    "<v>40</v>"
                  "</c>"
                "</row>"
              "</sheetData>"
              "<conditionalFormatting sqref=\"A1:A4\">"
                "<cfRule type=\"beginsWith\" priority=\"1\" operator=\"beginsWith\" text=\"b\">"
                  "<formula>LEFT(A1,1)=\"b\"</formula>"
                "</cfRule>"
                "<cfRule type=\"beginsWith\" priority=\"2\" operator=\"beginsWith\" text=\"bc\">"
                  "<formula>LEFT(A1,2)=\"bc\"</formula>"
                "</cfRule>"
                "<cfRule type=\"endsWith\" priority=\"3\" operator=\"endsWith\" text=\"z\">"
                  "<formula>RIGHT(A1,1)=\"z\"</formula>"
                "</cfRule>"
                "<cfRule type=\"endsWith\" priority=\"4\" operator=\"endsWith\" text=\"yz\">"
                  "<formula>RIGHT(A1,2)=\"yz\"</formula>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
            "</worksheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    worksheet_write_number(worksheet, CELL("A1"), 10, NULL);
    worksheet_write_number(worksheet, CELL("A2"), 20, NULL);
    worksheet_write_number(worksheet, CELL("A3"), 30, NULL);
    worksheet_write_number(worksheet, CELL("A4"), 40, NULL);

    lxw_conditional_format *conditional_format = calloc(1, sizeof(lxw_conditional_format));

    conditional_format->type         = LXW_CONDITIONAL_TYPE_TEXT;
    conditional_format->criteria     = LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH;
    conditional_format->value_string = "b";
    worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXW_CONDITIONAL_TYPE_TEXT;
    conditional_format->criteria     = LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH;
    conditional_format->value_string = "bc";
    worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXW_CONDITIONAL_TYPE_TEXT;
    conditional_format->criteria     = LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH;
    conditional_format->value_string = "z";
    worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXW_CONDITIONAL_TYPE_TEXT;
    conditional_format->criteria     = LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH;
    conditional_format->value_string = "yz";
    worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    free(conditional_format);

    lxw_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_worksheet_free(worksheet);
}
