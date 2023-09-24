/*****************************************************************************
 * format - A library for creating Excel XLSX format files.
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
#include "xlsxwriter/format.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"
#include "hbapierr.h"
#include "hbapiitm.h"


static HB_GARBAGE_FUNC( XLSXFormat_release )
{
	// printf( "Chiamato hb_XLSXFormat_release 2\n" );
   void ** ph = ( void ** ) Cargo;

   /* Check if pointer is not NULL to avoid multiple freeing */
   if( ph && *ph )
   {
      /* Destroy the object */
	 // printf( "Chiamato hb_XLSXFormat_release 3a\n" );
      lxw_format_free( ( lxw_format * ) *ph );
	 // printf( "Chiamato hb_XLSXFormat_release 3b\n" );

      /* set pointer to NULL to avoid multiple freeing */
      *ph = NULL;
   }
}

static const HB_GC_FUNCS s_gcXLSXFormatFuncs =
{
   XLSXFormat_release,
   hb_gcDummyMark
};

void hb_XLSXFormat_ret( lxw_format * p )
{
   
   if( p )
   {
      void ** ph = ( void ** ) hb_gcAllocate( sizeof( lxw_format * ), &s_gcXLSXFormatFuncs );

      *ph = p;

      hb_retptrGC( ph );
   }
   else
      hb_retptr( NULL );
}

lxw_format * hb_XLSXFormat_par( int iParam )
{
   void ** ph = ( void ** ) hb_parptrGC( &s_gcXLSXFormatFuncs, iParam );

   return ph ? ( lxw_format * ) *ph : NULL;
}

lxw_format * hb_XLSXFormat_item( PHB_ITEM pValue )
{
   void ** ph = ( void ** ) hb_itemGetPtrGC( pValue, &s_gcXLSXFormatFuncs );

   return ph ? ( lxw_format * ) *ph : NULL;
}

/*
 * Create a new format object.
 *
 * lxw_format *
 * lxw_format_new(void)
 *
 */
HB_FUNC( LXW_FORMAT_NEW )
{
	// printf( "Chiamato lxw_format_new\n" );
   lxw_format *format = lxw_format_new();
   hb_XLSXFormat_ret( format );
   // hb_retptr( format ); 
}




/*
 * Free a format object.
 *
 * void
 * lxw_format_free(lxw_format *format)
 * 
 */
HB_FUNC( LXW_FORMAT_FREE )
{  
	printf( "Chiamato LXW_FORMAT_FREE\n" );
   // lxw_format *format = hb_parptr( 1 ) ;

   // lxw_format_free( format ); 
}




/*
 * Check a user input color.
 *
 * lxw_color_t
 * lxw_format_check_color(lxw_color_t color)
 *
 */
/*
HB_FUNC( LXW_FORMAT_CHECK_COLOR )
{ 
   lxw_color_t color = hb_parnl( 1 ) ;

   hb_retnl( lxw_format_check_color( color ) ); 
}
*/





/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/






/*
 * Returns a font struct suitable for hashing as a lookup key.
 *
 * lxw_font *
 * lxw_format_get_font_key(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_FONT_KEY )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;

   if( self )
   hb_retptr( lxw_format_get_font_key( self ) ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Returns a border struct suitable for hashing as a lookup key.
 *
 * lxw_border *
 * lxw_format_get_border_key(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_BORDER_KEY )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   hb_retptr( lxw_format_get_border_key( self ) ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Returns a pattern fill struct suitable for hashing as a lookup key.
 *
 * lxw_fill *
 * lxw_format_get_fill_key(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_FILL_KEY )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   hb_retptr( lxw_format_get_fill_key( self ) ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Returns the XF index number used by Excel to identify a format.
 *
 * int32_t
 * lxw_format_get_xf_index(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_XF_INDEX )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;

   if( self )
   hb_retnl( lxw_format_get_xf_index( self ) ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}



/*
 * Set the font_name property.
 *
 * void
 * format_set_font_name(lxw_format *self, const char *font_name)
 *
 */
HB_FUNC( FORMAT_SET_FONT_NAME )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   const char *font_name = hb_parcx( 2 ) ;

   if( self )
   format_set_font_name( self, font_name ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_size property.
 *
 * void
 * format_set_font_size(lxw_format *self, double size)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SIZE )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   double size = hb_parnd( 2 ) ;

   if( self )
   format_set_font_size( self, size ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_color property.
 *
 * void
 * format_set_font_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_FONT_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

//   self->font_color = lxw_format_check_color(color);
   if( self )
   self->font_color = color;
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the bold property.
 *
 * void
 * format_set_bold(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_BOLD )
{ 
    // printf(" set bold\n" );

   lxw_format * self = hb_XLSXFormat_par( 1 ); // hb_parptr( 1 ) ;

   if ( self ) {
        // printf( "bold settato\n" );
       self->bold = LXW_TRUE;
   }
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
    // else
        // printf( "bold non settato\n" );
}




/*
 * Set the italic property.
 *
 * void
 * format_set_italic(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_ITALIC )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   self->italic = LXW_TRUE;
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the underline property.
 *
 * void
 * format_set_underline(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_UNDERLINE )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_underline( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_strikeout property.
 *
 * void
 * format_set_font_strikeout(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_STRIKEOUT )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   format_set_font_strikeout( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_script property.
 *
 * void
 * format_set_font_script(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SCRIPT )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_font_script( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_outline property.
 *
 * void
 * format_set_font_outline(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_OUTLINE )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;

   if( self )
   format_set_font_outline( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_shadow property.
 *
 * void
 * format_set_font_shadow(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SHADOW )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   format_set_font_shadow( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the num_format property.
 *
 * void
 * format_set_num_format(lxw_format *self, const char *num_format)
 *
 */
HB_FUNC( FORMAT_SET_NUM_FORMAT )
{ 
   lxw_format * self = hb_XLSXFormat_par( 1 ) ; // hb_parptr( 1 ) ;
   const char *num_format = hb_parcx( 2 ) ;

   if( self )
   format_set_num_format( self, num_format ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the unlocked property.
 *
 * void
 * format_set_unlocked(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_UNLOCKED )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   format_set_unlocked( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the hidden property.
 *
 * void
 * format_set_hidden(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_HIDDEN )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   format_set_hidden( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the align property.
 *
 * void
 * format_set_align(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_ALIGN )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;
   // fprintf( stderr, "sono qui\n" );

   if( self )
      format_set_align( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the text_wrap property.
 *
 * void
 * format_set_text_wrap(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_TEXT_WRAP )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;

   if( self )
   format_set_text_wrap( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the rotation property.
 *
 * void
 * format_set_rotation(lxw_format *self, int16_t angle)
 *
 */
HB_FUNC( FORMAT_SET_ROTATION )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   int16_t angle = hb_parnl( 2 ) ;

   if( self )
   format_set_rotation( self, angle ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the indent property.
 *
 * void
 * format_set_indent(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_INDENT )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   if( self )
   format_set_indent( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the shrink property.
 *
 * void
 * format_set_shrink(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_SHRINK )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   format_set_shrink( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the text_justlast property.
 *
 * void
 * format_set_text_justlast(lxw_format *self)
 *
 */
/*
HB_FUNC( FORMAT_SET_TEXT_JUSTLAST )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   format_set_text_justlast( self ); 
}
*/



/*
 * Set the pattern property.
 *
 * void
 * format_set_pattern(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_PATTERN )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   if( self )
   format_set_pattern( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the bg_color property.
 *
 * void
 * format_set_bg_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_BG_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   if( self )
   format_set_bg_color( self, color ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the fg_color property.
 *
 * void
 * format_set_fg_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_FG_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   if( self )
   format_set_fg_color( self, color ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the border property.
 *
 * void
 * format_set_border(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_BORDER )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_border( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the border_color property.
 *
 * void
 * format_set_border_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_BORDER_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   if( self )
   format_set_border_color( self, color ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the bottom property.
 *
 * void
 * format_set_bottom(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_BOTTOM )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_bottom( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the bottom_color property.
 *
 * void
 * format_set_bottom_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_BOTTOM_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   if( self )
   format_set_bottom_color( self, color) ; 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the left property.
 *
 * void
 * format_set_left(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_LEFT )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_left( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the left_color property.
 *
 * void
 * format_set_left_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_LEFT_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   if( self )
   format_set_left_color( self, color ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the right property.
 *
 * void
 * format_set_right(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_RIGHT )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_right( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the right_color property.
 *
 * void
 * format_set_right_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_RIGHT_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par(1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

   if( self )
   format_set_right_color( self, color ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the top property.
 *
 * void
 * format_set_top(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_TOP )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_top( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the top_color property.
 *
 * void
 * format_set_top_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_TOP_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

   if( self )
   format_set_top_color( self, color ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the diag_type property.
 *
 * void
 * format_set_diag_type(lxw_format *self, uint8_t type)
 *
 */
HB_FUNC( FORMAT_SET_DIAG_TYPE )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;

   if( self )
   format_set_diag_type( self, type ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the diag_color property.
 *
 * void
 * format_set_diag_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_DIAG_COLOR )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

   if( self )
   format_set_diag_color( self, color ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the diag_border property.
 *
 * void
 * format_set_diag_border(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_DIAG_BORDER )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   if( self )
   format_set_diag_border( self, style ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the num_format_index property.
 *
 * void
 * format_set_num_format_index(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_NUM_FORMAT_INDEX )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   if( self )
   format_set_num_format_index( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the valign property.
 *
 * void
 * format_set_valign(lxw_format *self, uint8_t value)
 *
 */
/*
HB_FUNC( FORMAT_SET_VALIGN )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   if( self )
   format_set_valign( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}
*/



/*
 * Set the reading_order property.
 *
 * void
 * format_set_reading_order(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_READING_ORDER )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   if( self )
   format_set_reading_order( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_family property.
 *
 * void
 * format_set_font_family(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_FONT_FAMILY )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   if( self )
   format_set_font_family( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_charset property.
 *
 * void
 * format_set_font_charset(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_FONT_CHARSET )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   if( self )
   format_set_font_charset( self, value ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_scheme property.
 *
 * void
 * format_set_font_scheme(lxw_format *self, const char *font_scheme)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SCHEME )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;
   const char *font_scheme = hb_parcx( 2 ) ;

   if( self )
   format_set_font_scheme( self, font_scheme ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_condense property.
 *
 * void
 * format_set_font_condense(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_CONDENSE )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   format_set_font_condense( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}




/*
 * Set the font_extend property.
 *
 * void
 * format_set_font_extend(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_EXTEND )
{ 
   lxw_format *self = hb_XLSXFormat_par( 1 ) ;

   if( self )
   format_set_font_extend( self ); 
   else
      hb_errRT_BASE( EG_ARG, 2020, NULL, HB_ERR_FUNCNAME, HB_ERR_ARGS_BASEPARAMS );
}


//eof

/*
 * Check a user input color.
 */
 /*
lxw_color_t lxw_format_check_color(lxw_color_t color)
{
    if (color == LXW_COLOR_UNSET)
        return color;
    else
        return color & LXW_COLOR_MASK;
}
*/
