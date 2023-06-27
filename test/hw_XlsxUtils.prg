FUNCTION hw_XlsxExport( aSheet, aOptions )

   LOCAL oWorkBook, oWorkSheet, oTitleFormat, aRows := {}, aHeader := {}, aRow := {}, aCol := {}, hRow := { => }, aColWidth := {}, aTitle := {}, aFooter := {}
   LOCAL hOptions := fillOptions( aOptions )
   LOCAL nRow := 0, nCol := 0

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
      oTitleFormat := workbook_add_format( oWorkBook )
      format_set_bold( oTitleFormat )
      IF ValType( hOptions[ 'Title' ] ) == "A"
         FOR EACH aTitle in hOptions[ 'Title' ]
            worksheet_write_string( oWorkSheet, nRow, 0, aTitle, oTitleFormat )
            nRow++
         NEXT
      ELSE
         worksheet_write_string( oWorkSheet, nRow, 0, hOptions[ 'Title' ], oTitleFormat )
         nRow++
      ENDIF
   ENDIF

   IF !Empty( aHeader )
      FOR EACH aRow in aHeader
         worksheet_write_string( oWorkSheet, nRow, nCol, aRow )
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
            worksheet_write_string( oWorkSheet, nRow, nCol, hRow[ aCol ] )
            IF Empty( aColWidth[ nCol + 1 ] )
               aColWidth[ nCol + 1 ] := 0
            ENDIF
            IF aColWidth[ nCol + 1 ] < Len( AllTrim( hRow[ aCol ] ) ) * 1.2
               aColWidth[ nCol + 1 ] := Len( AllTrim( hRow[ aCol ] ) ) * 1.2
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
            worksheet_write_string( oWorkSheet, nRow, 0, aFooter, oFooterFormat )
            nRow++
         NEXT
      ELSE
         worksheet_write_string( oWorkSheet, nRow, 0, hOptions[ 'Footer' ], oFooterFormat )
         nRow++
      ENDIF
   ENDIF

   workbook_close( oWorkBook )

RETURN

static FUNCTION fillOptions( aOptions )

   LOCAL aOption := {}, hOptions := { => }

   hb_HCaseMatch( hOptions, .F. )
   hOptions := { => }
   hOptions[ "FileName" ]  := "libro.xlsx"
   hOptions[ "SheetName" ] := "hoja"
   hOptions[ "Headers" ]   := {}
   hOptions[ "Title" ]     := {}
   hOptions[ "Footer" ]    := {}

   FOR EACH aOption in aOptions
      hOptions[ aOption:__enumkey() ] := aOption
   NEXT

RETURN hOptions

static FUNCTION array2Hash( aArray )

   LOCAL hHash := { => }, aElement := {}

   hb_HCaseMatch( hHash, .F. )

   FOR EACH aElement in aArray
      hHash[ aElement:__enumkey() ] := aElement
   NEXT

RETURN hHash