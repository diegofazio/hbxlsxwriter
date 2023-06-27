FUNCTION hw_XlsxExport( aSheet, aOptions )

   LOCAL oWorkBook, oWorkSheet, oTitleFormat1, oTitleFormat2, aRows := {}, aHeader := {}, aRow := {}, aCol := {}, hRow := { => }, aColWidth := {}, aTitle := {}, aFooter := {}
   LOCAL hOptions := fillOptions( aOptions ), oFormatHeader
   LOCAL nRow := 0, nCol := 0, cStr := ""

   FOR EACH aRow in aSheet[ 1 ]
      IF Empty( hOptions[ 'aHeaders' ] )
         AAdd( aHeader, aRow:__enumkey() )
      ELSE
         IF ( hb_AScan( hOptions[ 'aHeaders' ], aRow:__enumkey(),, .F. ) > 0 )
            AAdd( aHeader, aRow:__enumkey() )
         ENDIF
      ENDIF
   NEXT

   oWorkBook  := workbook_new( hOptions[ 'FileName' ] )
   oWorkSheet := workbook_add_worksheet( oWorkBook, hOptions[ 'SheetName' ] )

   IF !Empty( hOptions[ 'Title' ] )
      oTitleFormat1 := workbook_add_format( oWorkBook )
      format_set_bold( oTitleFormat1 )
      format_set_align( oTitleFormat1, LXW_ALIGN_CENTER )
      format_set_align( oTitleFormat1, LXW_ALIGN_VERTICAL_CENTER )
      format_set_font_size( oTitleFormat1, 14 )
      IF ValType( hOptions[ 'Title' ] ) == "A"
         oTitleFormat2 := workbook_add_format( oWorkBook )
         format_set_font_size( oTitleFormat2, 12 )
         format_set_bold( oTitleFormat2 )
         format_set_align( oTitleFormat2, LXW_ALIGN_LEFT )
         format_set_align( oTitleFormat2, LXW_ALIGN_VERTICAL_CENTER )
         FOR EACH aTitle in hOptions[ 'Title' ]
            IF nRow == 0
               worksheet_merge_range( oWorkSheet, nRow, 0, nRow, Len( aHeader ) - 1, aTitle, oTitleFormat1 )
            ELSE
               worksheet_merge_range( oWorkSheet, nRow, 0, nRow, Len( aHeader ) - 1, aTitle, oTitleFormat2 )
            ENDIF
            nRow++
         NEXT
      ELSE
         worksheet_merge_range( oWorkSheet, nRow, 0, nRow, Len( aHeader ) - 1, hOptions[ 'Title' ], oTitleFormat2 )
         nRow++
      ENDIF
   ENDIF

   IF !Empty( aHeader )

      oFormatHeader := workbook_add_format( oWorkBook )
      format_set_bottom( oFormatHeader, LXW_BORDER_THIN )
      format_set_top( oFormatHeader, LXW_BORDER_THIN )
      format_set_left( oFormatHeader, LXW_BORDER_THIN )
      format_set_right( oFormatHeader, LXW_BORDER_THIN )
      format_set_bottom_color( oFormatHeader, 0x808080 )
      format_set_top_color( oFormatHeader, 0x808080 )
      format_set_left_color( oFormatHeader, 0x808080 )
      format_set_right_color( oFormatHeader, 0x808080 )
      format_set_bg_color( oFormatHeader, 0xd1d1d1 )
      format_set_pattern( oFormatHeader, LXW_PATTERN_SOLID )

      FOR EACH aRow in aHeader
         worksheet_write_string( oWorkSheet, nRow, nCol, aRow, oFormatHeader )
         nCol++
      NEXT
      nRow++
   ENDIF

   FOR EACH aCol in aHeader
      AAdd( aColWidth, Len( AllTrim( aCol ) ) * 1.2 )
   NEXT

   FOR EACH aRow in aSheet
      nCol := 0
      hRow := array2Hash( aRow )
      FOR EACH aCol in aHeader
         IF ValType( hRow[ aCol ] ) == "N"
            worksheet_write_number( oWorkSheet, nRow, nCol, hRow[ aCol ] )
            IF Empty( aColWidth[ nCol + 1 ] )
               aColWidth[ nCol + 1 ] := 0
            ENDIF
            IF aColWidth[ nCol + 1 ] < Len( AllTrim( Str( hRow[ aCol ] ) ) ) * 1.2
               aColWidth[ nCol + 1 ] := Len( AllTrim( Str( hRow[ aCol ] ) ) ) * 1.2
            ENDIF
            worksheet_set_column( oWorkSheet, nRow, nCol, aColWidth[ nCol + 1 ] )
         ELSE
            cStr := hw_ValtoChar( hRow[ aCol ] )
            worksheet_write_string( oWorkSheet, nRow, nCol, cStr )
            IF Empty( aColWidth[ nCol + 1 ] )
               aColWidth[ nCol + 1 ] := 0
            ENDIF
            IF aColWidth[ nCol + 1 ] < Len( AllTrim( cStr ) ) * 1.2
               aColWidth[ nCol + 1 ] := Len( AllTrim( cStr ) ) * 1.2
            ENDIF
            worksheet_set_column( oWorkSheet, nRow, nCol, aColWidth[ nCol + 1 ] )
         ENDIF
         nCol++
      NEXT
      nRow++
   NEXT

   IF !Empty( hOptions[ 'Footer' ] )
      oFooterFormat := workbook_add_format( oWorkBook )
      format_set_bold( oFooterFormat )
      IF ValType( hOptions[ 'Footer' ] ) == "A"
         FOR EACH aFooter in hOptions[ 'Footer' ]
            worksheet_merge_range( oWorkSheet, nRow, 0, nRow, Len( aHeader ) - 1, aFooter, oFooterFormat )
            nRow++
         NEXT
      ELSE
         worksheet_merge_range( oWorkSheet, nRow, 0, nRow, Len( aHeader ) - 1, hOptions[ 'Footer' ], oFooterFormat )
         nRow++
      ENDIF
   ENDIF

   workbook_close( oWorkBook )

RETURN iif( hOptions[ 'ReturnContent' ], hb_MemoRead( hOptions[ 'FileName' ] ), NIL )

STATIC FUNCTION fillOptions( aOptions )

   LOCAL aOption := {}, hOptions := { => }
   LOCAL cTempFile := AllTrim( GetEnv( "TEMP" ) ) + "\xlsx" + StrZero( hb_RandomInt( 1, 10 ^ ( 8 - Len( 'xlsx' ) ) - 1 ), 8 - Len( 'xlsx' ), 0 ) + ".xlsx"

   hb_HCaseMatch( hOptions, .F. )
   hOptions := { => }
   hOptions[ "FileName" ]        := cTempFile
   hOptions[ "SheetName" ]       := "hoja"
   hOptions[ "Headers" ]         := {}
   hOptions[ "Title" ]           := {}
   hOptions[ "Footer" ]          := {}
   hOptions[ "ReturnContent" ]   := .F.

   FOR EACH aOption in aOptions
      hOptions[ aOption:__enumkey() ] := aOption
   NEXT

RETURN hOptions

STATIC FUNCTION array2Hash( aArray )

   LOCAL hHash := { => }, aElement := {}

   hb_HCaseMatch( hHash, .F. )

   FOR EACH aElement in aArray
      hHash[ aElement:__enumkey() ] := aElement
   NEXT

RETURN hHash
