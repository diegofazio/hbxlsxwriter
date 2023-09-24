/*****************************************************************************
 * workbook - A library for creating Excel XLSX workbook files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2019, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */
/*
 * Wrapped for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/workbook.h"
#include "xlsxwriter/utility.h"
#include "xlsxwriter/packager.h"
#include "xlsxwriter/hash_table.h"


#include "hbapierr.h"
#include "hbapiitm.h"


static HB_GARBAGE_FUNC( XLSXWorkbook_release )
{
	// printf( "Chiamato hb_XLSXWorkbook_release 2\n" );
   void ** ph = ( void ** ) Cargo;

   /* Check if pointer is not NULL to avoid multiple freeing */
   if( ph && *ph )
   {
      /* Destroy the object */
	 // printf( "Chiamato hb_XLSXWorkbook_release 3a\n" );
      lxw_workbook_free( ( lxw_workbook * ) *ph );
	 // printf( "Chiamato hb_XLSXWorkbook_release 3b\n" );

      /* set pointer to NULL to avoid multiple freeing */
      *ph = NULL;
   }
}

static const HB_GC_FUNCS s_gcXLSXWorkbookFuncs =
{
   XLSXWorkbook_release,
   hb_gcDummyMark
};

void hb_XLSXWorkbook_ret( lxw_workbook * p )
{
    // fprintf( stderr,"Chiamato hb_XLSXWorkbook_ret\n" );
   if( p )
   {
      void ** ph = ( void ** ) hb_gcAllocate( sizeof( lxw_workbook * ), &s_gcXLSXWorkbookFuncs );

      *ph = p;

      hb_retptrGC( ph );
   }
   else
      hb_retptr( NULL );
}

lxw_workbook * hb_XLSXWorkbook_par( int iParam )
{
   void ** ph = ( void ** ) hb_parptrGC( &s_gcXLSXWorkbookFuncs, iParam );

   return ph ? ( lxw_workbook * ) *ph : NULL;
}

lxw_workbook * hb_XLSXWorkbook_item( PHB_ITEM pValue )
{
   void ** ph = ( void ** ) hb_itemGetPtrGC( pValue, &s_gcXLSXWorkbookFuncs );

   return ph ? ( lxw_workbook * ) *ph : NULL;
}

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Free a workbook object.
 *
 * void
 * lxw_workbook_free(lxw_workbook *workbook)
 *
 */
HB_FUNC( LXW_WORKBOOK_FREE )
{
       //	
       printf( "LXW_WORKBOOK_FREE non deve essere chiamata direttamente\n");
   //lxw_workbook *workbook = hb_XLSXWorkbook_par( 1 ) ;

   //lxw_workbook_free( workbook ); 
}




/*
 * Set the default index for each format. This is only used for testing.
 *
 * void
 * lxw_workbook_set_default_xf_indices(lxw_workbook *self)
 *
 */
HB_FUNC( LXW_WORKBOOK_SET_DEFAULT_XF_INDICES )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;

   lxw_workbook_set_default_xf_indices( self ); 
}





/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/


/*
 * Assemble and write the XML file.
 *
 * void
 * lxw_workbook_assemble_xml_file(lxw_workbook *self)
 *
 */
HB_FUNC( LXW_WORKBOOK_ASSEMBLE_XML_FILE )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;

   lxw_workbook_assemble_xml_file( self ); 
}




/*
 *
 * Public functions.
 *
 ****************************************************************************/





/*
 * Create a new workbook object.
 *
 * lxw_workbook *
 * workbook_new(const char *filename)
 *
 */
HB_FUNC( WORKBOOK_NEW )
{ 
   const char *filename = hb_parcx( 1 ) ;

   hb_XLSXWorkbook_ret( workbook_new( filename ) ); 
}

/* Deprecated function name for backwards compatibility. */
/*
lxw_workbook *
new_workbook(const char *filename)
*/
HB_FUNC( NEW_WORKBOOK )
{
   const char *filename = hb_parcx( 1 ) ;
   hb_XLSXWorkbook_ret( workbook_new_opt(filename, NULL) );
}


/*
 * Create a new workbook object with options.
 *
 * lxw_workbook *
 * workbook_new_opt(const char *filename, lxw_workbook_options *options)
 *
 */
HB_FUNC( WORKBOOK_NEW_OPT )
{
   const char *filename = hb_parcx( 1 );
   lxw_workbook_options *options = hb_param( 2, HB_IT_ANY );
   if HB_ISNIL( 2 )
   {
      hb_XLSXWorkbook_ret( workbook_new_opt(filename, NULL));
   }
   else
   {
      hb_XLSXWorkbook_ret(workbook_new_opt(filename, options));
   }
}




/*
 * Add a new worksheet to the Excel workbook.
 *
 * lxw_worksheet *
 * workbook_add_worksheet(lxw_workbook *self, const char *sheetname)
 *
 */
HB_FUNC( WORKBOOK_ADD_WORKSHEET )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par(1);
   const char *sheetname = hb_parcx( 2 );
   if ( HB_ISNIL( 2 ) || strlen(sheetname) == 0 )
   {
      hb_retptr( workbook_add_worksheet( self, NULL ) );
   }
   else
   {
      hb_retptr( workbook_add_worksheet( self, sheetname ) );
   }
}




/*
 * Add a new chartsheet to the Excel workbook.
 *
 * lxw_chartsheet *
 * workbook_add_chartsheet(lxw_workbook *self, const char *sheetname)
 *
 */
HB_FUNC( WORKBOOK_ADD_CHARTSHEET )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *sheetname = hb_parcx( 2 ) ;

   hb_retptr( workbook_add_chartsheet( self, sheetname ) ); 
}




/*
 * Add a new chart to the Excel workbook.
 *
 * lxw_chart *
 * workbook_add_chart(lxw_workbook *self, uint8_t type)
 *
 */
HB_FUNC( WORKBOOK_ADD_CHART )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;

   hb_retptr( workbook_add_chart( self, type ) ); 
}

void hb_XLSXFormat_ret( lxw_format * p ) ;


/*
 * Add a new format to the Excel workbook.
 *
 * lxw_format *
 * workbook_add_format(lxw_workbook *self)
 *
 */
HB_FUNC( WORKBOOK_ADD_FORMAT )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;

   hb_XLSXFormat_ret( workbook_add_format( self ) ); 
   // hb_retptr( workbook_add_format( self ) ); 
}




/*
 * Call finalization code and close file.
 *
 * lxw_error
 * workbook_close(lxw_workbook *self)
 *
 */
HB_FUNC( WORKBOOK_CLOSE )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;

   hb_retni( workbook_close( self ) ); 
}




/*
 * Create a defined name in Excel. We handle global/workbook level names and
 * local/worksheet names.
 *
 * lxw_error
 * workbook_define_name(lxw_workbook *self, const char *name,
 *    const char *formula)
 *
 */
HB_FUNC( WORKBOOK_DEFINE_NAME )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   const char *formula = hb_parcx( 3 ) ;

   hb_retni( workbook_define_name( self, name, formula ) ); 
}


/*
 *  Code borrowed and adapted
 *
 *  lxw takes care of copying and safely store the strings
 */
lxw_doc_properties * hash2properties( PHB_ITEM pHash )
{
   if( HB_IS_HASH( pHash ) )
   {
      lxw_doc_properties *properties = (lxw_doc_properties *) hb_xalloc( sizeof(lxw_doc_properties) ); 
 
      memset( properties, 0, sizeof( lxw_doc_properties ) );

      HB_SIZE nLen = hb_hashLen( pHash ), nPos = 0;

      while( ++nPos <= nLen )
      {
         PHB_ITEM pKey = hb_hashGetKeyAt( pHash, nPos );
         PHB_ITEM pValue = hb_hashGetValueAt( pHash, nPos );
         if( pKey && pValue )
         {
            char *key = (char *)hb_itemGetC( pKey );

            if( HB_IS_STRING( pValue ) )
            {
                char *value = (char *) hb_itemGetC( pValue );

                if( hb_stricmp( key, "title" ) == 0 ){
                   properties->title = value;
                }
                else if( hb_stricmp( key, "subject" ) == 0 ){
                   properties->subject = value;
                }
                else if( hb_stricmp( key, "author" ) == 0 ){
                   properties->author = value;
                }
                else if( hb_stricmp( key, "manager" ) == 0 ){
                   properties->manager = value;
                }
                else if( hb_stricmp( key, "company" ) == 0 ){
                   properties->company = value;
                }
                else if( hb_stricmp( key, "category" ) == 0 ){
                   properties->category = value;
                }
                else if( hb_stricmp( key, "keywords" ) == 0 ){
                   properties->keywords = value;
                }
                else if( hb_stricmp( key, "comments" ) == 0 ){
                   properties->comments = value;
                }
                else if( hb_stricmp( key, "status" ) == 0 ){
                   properties->status = value;
                }
                else if( hb_stricmp( key, "hyperlink_base" ) == 0 ){
                   properties->hyperlink_base = value;
                }
            }
	 }
      }
      if( properties ){
         return properties; 
      }
   }
   return 0;
}


/*
 * Set the document properties such as Title, Author etc.
 *
 * lxw_error
 * workbook_set_properties(lxw_workbook *self, lxw_doc_properties *user_props)
 *
 */
HB_FUNC( WORKBOOK_SET_PROPERTIES )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   PHB_ITEM pHash = hb_param( 2, HB_IT_HASH );

   lxw_doc_properties *user_props = hash2properties( pHash ) ;

   hb_retni( workbook_set_properties( self, user_props ) ); 

   hb_xfree( user_props );
}




/*
 * Set a string custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_string(lxw_workbook *self, const char *name,
 *      const char *value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_STRING )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   const char *value = hb_parcx( 3 ) ;

   hb_retni( workbook_set_custom_property_string( self, name, value ) ); 
}




/*
 * Set a double number custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_number(lxw_workbook *self, const char *name,
 *       double value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_NUMBER )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   double value = hb_parnd( 3 ) ;

   hb_retni( workbook_set_custom_property_number( self, name, value ) ); 
}




/*
 * Set a integer number custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_integer(lxw_workbook *self, const char *name,
 *        int32_t value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_INTEGER )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   int32_t value = hb_parnl(3 ) ;

   hb_retni( workbook_set_custom_property_integer( self, name, value ) ); 
}




/*
 * Set a boolean custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_boolean(lxw_workbook *self, const char *name,
 *          uint8_t value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_BOOLEAN )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   uint8_t value = hb_parni( 3 ) ;

   hb_retni( workbook_set_custom_property_boolean( self, name, value ) ); 
}




/*
 * Set a datetime custom document property.
 *
 * lxw_error 
 * workbook_set_custom_property_datetime(lxw_workbook *self, const char *name,
 *           lxw_datetime *datetime)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_DATETIME )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   lxw_datetime *datetime = hb_parptr(3 ) ;

   hb_retni( workbook_set_custom_property_datetime( self, name, datetime ) ); 
}




/*
 * Get a worksheet object from its name.
 *
 * lxw_worksheet *
 * workbook_get_worksheet_by_name(lxw_workbook *self, const char *name)
 *
 */
HB_FUNC( WORKBOOK_GET_WORKSHEET_BY_NAME )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   hb_retptr( workbook_get_worksheet_by_name( self, name ) ); 
}




/*
 * Get a chartsheet object from its name.
 *
 * lxw_chartsheet *
 * workbook_get_chartsheet_by_name(lxw_workbook *self, const char *name)
 *
 */
HB_FUNC( WORKBOOK_GET_CHARTSHEET_BY_NAME )
{ 
   lxw_workbook *self = hb_XLSXWorkbook_par( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   hb_retptr( workbook_get_chartsheet_by_name( self, name ) ); 
}




//eof
