/*
 * Create a worksheet with a chart and implement customs colors
 *
 * by Fazio Diego, diegohfazio@gmail.com
 */

#include "hbxlsxwriter.ch"

FUNCTION Main()

   LOCAL workbook, worksheet, chart
   LOCAL pFillRed
   LOCAL pFillGreen
   LOCAL pFillBlue
   LOCAL pP11, pP12, pP13
   LOCAL pP21, pP22, pP23
   LOCAL pP31, pP32, pP33

   workbook  = new_workbook( "chart.xlsx" )
   worksheet = workbook_add_worksheet( workbook, NIL )

   /* Write some data for the chart. */
   write_worksheet_data( worksheet )

   /* Create a chart object. */
   chart = workbook_add_chart( workbook, LXW_CHART_COLUMN )

    /* Configure the chart. In simplest case we just add some value data
     * series. The NULL categories will default to 1 to 5 like in Excel.
     */
   oSeries1 := chart_add_series( chart, NIL, "Sheet1!$B$2:$B$4" )
   oSeries2 := chart_add_series( chart, NIL, "Sheet1!$C$2:$C$4" )
   oSeries3 := chart_add_series( chart, NIL, "Sheet1!$D$2:$D$4" )

   pFillRed   := Chart_Fill_New( LXW_COLOR_BLACK )
   pFillGreen := Chart_Fill_New( LXW_COLOR_GREEN )
   pFillBlue  := Chart_Fill_New( LXW_COLOR_RED )

   /* Points for each serie (three points for serie) */
   pP11 := Chart_Point_New()
   pP12 := Chart_Point_New()
   pP13 := Chart_Point_New()

   pP21 := Chart_Point_New()
   pP22 := Chart_Point_New()
   pP23 := Chart_Point_New()

   pP31 := Chart_Point_New()
   pP32 := Chart_Point_New()
   pP33 := Chart_Point_New()

   /* Assign colors: serie1 = red, serie2 = green, serie3 = blue */
   Chart_Point_Set_Fill( pP11, pFillRed )
   Chart_Point_Set_Fill( pP12, pFillRed )
   Chart_Point_Set_Fill( pP13, pFillRed )

   Chart_Point_Set_Fill( pP21, pFillGreen )
   Chart_Point_Set_Fill( pP22, pFillGreen )
   Chart_Point_Set_Fill( pP23, pFillGreen )

   Chart_Point_Set_Fill( pP31, pFillBlue )
   Chart_Point_Set_Fill( pP32, pFillBlue )
   Chart_Point_Set_Fill( pP33, pFillBlue )

   /* Apply points to each serie */

   Chart_Series_Set_Points( oSeries1, { pP11, pP12, pP13 } )
   Chart_Series_Set_Points( oSeries2, { pP21, pP22, pP23 } )
   Chart_Series_Set_Points( oSeries3, { pP31, pP32, pP33 } )

   /* Insert the chart into the worksheet. */
   worksheet_insert_chart( worksheet, LXW_CELL( "B7" ), chart )

RETURN workbook_close( workbook )

/* Write some data to the worksheet. */
PROCEDURE write_worksheet_data( worksheet )

   LOCAL aData, row, col

   aData := { ;
      { 1,  2,  3 }, ;
      { 2,  4,  6 }, ;
      { 3,  6,  9 }, ;
      }

   FOR row = 1 TO 3
      FOR col = 1 TO 3
         worksheet_write_number( worksheet, row, col, aData[ row ][ col ], NIL )
      NEXT col
   NEXT row

RETURN
