/*
 * Tests for the libxlsxwriter library.
 *
 * Copyright 2014-2024, John McNamara, jmcnamara@cpan.org
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "xlsxwriter/rich_value_types.h"

// Test _xml_declaration().
CTEST(rich_value_types, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = tmpfile();

    lxw_rich_value_types *rich_value_types = lxw_rich_value_types_new();
    rich_value_types->file = testfile;

    _rich_value_types_xml_declaration(rich_value_types);

    RUN_XLSX_STREQ(exp, got);

    lxw_rich_value_types_free(rich_value_types);
}
