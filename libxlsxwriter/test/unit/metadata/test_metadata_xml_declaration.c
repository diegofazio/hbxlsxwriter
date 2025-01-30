/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2024, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "xlsxwriter/metadata.h"

// Test _xml_declaration().
CTEST(metadata, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = tmpfile();

    lxw_metadata *metadata = lxw_metadata_new();
    metadata->file = testfile;

    _metadata_xml_declaration(metadata);

    RUN_XLSX_STREQ(exp, got);

    lxw_metadata_free(metadata);
}
