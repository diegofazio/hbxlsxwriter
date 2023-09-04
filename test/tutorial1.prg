/*
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 1 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "hbxlsxwriter.ch"

#define ITEM  1
#define COST  2

function main() 

    local workbook, worksheet, row, col, aExpenses

    aExpenses := { ;
	           { "Rent", 1000 },;
		   { "Gas",   100 },;
		   { "Food",  300 },;
		   { "Gym",    50 } ;
                 }

    /* Create a workbook and add a worksheet. */
    workbook  = workbook_new("tutorial01.xlsx")
    worksheet = workbook_add_worksheet(workbook, NIL)
    // Properties should be added as #defines to hbxlsxwriter.ch
    set_doc_property( 1, "This is the title" )
    set_doc_property( 1, "This is another title" )
    set_doc_property( 2, "This is the subject" )
    set_doc_property( 3, "This is the author" )
    set_doc_property( 4, "This is the manager" )
    set_doc_property( 5, "This is the company" )
    set_doc_property( 6, "This is the category" )
    set_doc_property( 7, "This is the keywords" )
    set_doc_property( 8, "This is the comments" )
    set_doc_property( 9, "This is the status" )
    set_doc_property(10, "This is the hyperlink" )

    col := 0

    /* Iterate over the data and write it out element by element. */
    for row := 1 to 4 
        worksheet_write_string(worksheet, row, col,     aExpenses[row][ITEM], NIL)
        worksheet_write_number(worksheet, row, col + 1, aExpenses[row][COST], NIL)
    next row

    /* Write a total using a formula. */
    worksheet_write_string (worksheet, row, col,     "Total",       NIL)
    worksheet_write_formula(worksheet, row, col + 1, "=SUM(B1:B4)", NIL)

    workbook_set_properties( workbook, NIL )

    /* Save the workbook and free any allocated memory. */
    return workbook_close(workbook)

//eof
