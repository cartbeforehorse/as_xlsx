CREATE OR REPLACE PACKAGE AS_XLSX IS
/*****************************************************************************
 *****************************************************************************
 **
 ** Author: Anton Scheffer
 ** Date: 19-02-2011
 ** Website: http://technology.amis.nl/blog
 ** See also: http://technology.amis.nl/blog/?p=10995
 **
 ** Changelog:
 **   Date: 21-02-2011
 **     Added Aligment, horizontal, vertical, wrapText
 **   Date: 06-03-2011
 **     Added Comments, MergeCells, fixed bug for dependency on NLS-settings
 **   Date: 16-03-2011
 **     Added bold and italic fonts
 **   Date: 22-03-2011
 **     Fixed issue with timezone's set to a region(name) instead of a offset
 **   Date: 08-04-2011
 **     Fixed issue with XML-escaping from text
 **   Date: 27-05-2011
 **     Added MIT-license
 **   Date: 11-08-2011
 **     Fixed NLS-issue with column width
 **   Date: 29-09-2011
 **     Added font color
 **   Date: 16-10-2011
 **     fixed bug in add_string
 **   Date: 26-04-2012
 **     Fixed set_autofilter (only one autofilter per sheet, added _xlnm._FilterDatabase)
 **     Added list_validation = drop-down
 **   Date: 27-08-2013
 **     Added freeze_pane
 **   Date: 05-09-2013
 **     Performance
 **   Date: 14-07-2014
 **      Added p_UseXf to query2sheet
 **   Date: 23-10-2014
 **      Added xml:space="preserve"
 **   Date: 29-02-2016
 **     Fixed issue with alignment in get_XfId
 **     Thank you Bertrand Gouraud
 **   Date: 01-04-2017
 **     Added p_height to set_row
 **   Date: 23-05-2018
 **     fixed bug in add_string (thank you David Short)
 **     added tabColor to new_sheet
 **   Date: 13-06-2018
 **     added  c_version
 **     added formulas
 **   Date: 12-02-2020
 **     added sys_refcursor overload of query2sheet
 **     use default date format in query2sheet
 **     changed to date1904=false
 *****************************************************************************
 *****************************************************************************

  Copyright (C) 2011, 2020 by Anton Scheffer

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in
  all copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
  THE SOFTWARE.

 *****************************************************************************
 *****************************************************************************/

--------------------------------------------------
-- Public Types
--
TYPE tp_alignment IS RECORD (
   vertical   VARCHAR2(11),
   horizontal VARCHAR2(16),
   wrapText   BOOLEAN );

TYPE data_binder IS RECORD (
   datatype  VARCHAR2(6), -- NUMBER,STRING,DATE
   s_val     VARCHAR2(2000),
   n_val     NUMBER,
   d_val     DATE );
TYPE bind_arr IS TABLE OF data_binder INDEX BY VARCHAR2(50);

TYPE param_rec IS RECORD (
   param_name      VARCHAR2(100),
   param_value     VARCHAR2(100),
   additional_info VARCHAR2(300) );
TYPE params_arr IS TABLE OF param_rec;

--------------------------------------------------
-- Fonts and fills stored by ID
--
TYPE fonts_list  IS TABLE OF INTEGER INDEX BY VARCHAR2(50);
TYPE fills_list  IS TABLE OF INTEGER INDEX BY VARCHAR2(50);
TYPE borddr_list IS TABLE OF INTEGER INDEX BY VARCHAR2(50);
TYPE numFmt_list IS TABLE OF INTEGER INDEX BY VARCHAR2(50);
TYPE align_list  IS TABLE OF tp_alignment INDEX BY VARCHAR2(50);
TYPE numFmt_cols IS TABLE OF INTEGER INDEX BY PLS_INTEGER;

fonts_  fonts_list;
fills_  fills_list;
bdrs_   borddr_list;
numFmt_ numFmt_list;
align_  align_list;

--------------------------------------------------
-- Public Procedures and Functions
--
PROCEDURE Init_Workbook;

PROCEDURE Clear_Workbook;

FUNCTION New_Sheet (
   sheetname_ VARCHAR2 := null,
   tab_color_ VARCHAR2 := null ) RETURN PLS_INTEGER;

PROCEDURE New_Sheet (
   sheetname_ VARCHAR2 := null,
   tab_color_ VARCHAR2 := null );

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 );

FUNCTION OraFmt2Excel (
   p_format IN VARCHAR2 := null ) RETURN VARCHAR2;

FUNCTION Get_NumFmt (
   format_ IN VARCHAR2 := null ) RETURN PLS_INTEGER;

PROCEDURE Set_Font (
   name_      IN VARCHAR2    := 'Calibri',
   sheet_     IN PLS_INTEGER := null,
   family_    IN PLS_INTEGER := 2,
   fontsize_  IN NUMBER      := 11,
   theme_     IN PLS_INTEGER := 1,
   underline_ IN BOOLEAN     := false,
   italic_    IN BOOLEAN     := false,
   bold_      IN BOOLEAN     := false,
   rgb_       IN VARCHAR2    := null ); -- hex Alpha-rgb value

FUNCTION Get_Font (
   name_      IN VARCHAR2    := 'Calibri',
   family_    IN PLS_INTEGER := 2,
   fontsize_  IN NUMBER      := 11,
   theme_     IN PLS_INTEGER := 1,
   underline_ IN BOOLEAN     := false,
   italic_    IN BOOLEAN     := false,
   bold_      IN BOOLEAN     := false,
   rgb_       IN VARCHAR2    := null ) RETURN PLS_INTEGER; -- hex Alpha-rgb value

FUNCTION Get_Fill (
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,                      -- hex Alpha-rgb value
   bgRGB_       IN VARCHAR2 := null ) RETURN PLS_INTEGER; -- hex Alpha-rgb value

PROCEDURE Get_Fill (
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,
   bgRGB_       IN VARCHAR2 := null );

---------------------------------------
-- Get_Border()
--  Values allowed in all these parameters are as follows:
--    none;thin;medium;dashed;dotted;thick;double;hair;mediumDashed;
--    dashDot;mediumDashDot;dashDotDot;mediumDashDotDot;slantDashDot
--
FUNCTION Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' ) RETURN PLS_INTEGER;

PROCEDURE Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' );

---------------------------------------
-- Get_Alignment()
--  Values allowed in vert/horiz: horizontal;center;centerContinuous;distributed;fill;general;justify;left;right
--  Values allowed in wrapText:   vertical;bottom;center;distributed;justify;top
--
FUNCTION Get_Alignment (
   vertical_   IN VARCHAR2 := null,
   horizontal_ IN VARCHAR2 := null,
   wrapText_   IN BOOLEAN  := null ) RETURN tp_alignment;

PROCEDURE Cell ( -- NUMBER
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN NUMBER,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null );
PROCEDURE CellP (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     NUMBER,
   numFmtId_  VARCHAR2    := null,
   fontId_    VARCHAR2    := null,
   fillId_    VARCHAR2    := null,
   borderId_  VARCHAR2    := null,
   alignment_ VARCHAR2    := null,
   sheet_     PLS_INTEGER := null );

PROCEDURE Cell ( -- VARCHAR
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN VARCHAR2,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null );
PROCEDURE CellP (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     VARCHAR2,
   numFmtId_  VARCHAR2    := null,
   fontId_    VARCHAR2    := null,
   fillId_    VARCHAR2    := null,
   borderId_  VARCHAR2    := null,
   alignment_ VARCHAR2    := null,
   sheet_     PLS_INTEGER := null );

PROCEDURE Cell ( -- DATE
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN DATE,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null );
PROCEDURE CellP (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     DATE,
   numFmtId_  VARCHAR2    := null,
   fontId_    VARCHAR2    := null,
   fillId_    VARCHAR2    := null,
   borderId_  VARCHAR2    := null,
   alignment_ VARCHAR2    := null,
   sheet_     PLS_INTEGER := null );

PROCEDURE Condition_Color_Col (
   col_   PLS_INTEGER,
   sheet_ PLS_INTEGER := null );

PROCEDURE Hyperlink (
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER,
   url_   IN VARCHAR2,
   value_ IN VARCHAR2    := null,
   sheet_ IN PLS_INTEGER := null );

PROCEDURE Comment (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   text_   IN VARCHAR2,
   author_ IN VARCHAR2 := null,
   width_  IN PLS_INTEGER := 150,  -- pixels
   height_ IN PLS_INTEGER := 100,  -- pixels
   sheet_  IN PLS_INTEGER := null );

PROCEDURE Num_Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN NUMBER       := null,
   numFmtId_      IN PLS_INTEGER  := null,
   fontId_        IN PLS_INTEGER  := null,
   fillId_        IN PLS_INTEGER  := null,
   borderId_      IN PLS_INTEGER  := null,
   alignment_     IN tp_alignment := null,
   sheet_         IN PLS_INTEGER  := null );

PROCEDURE Str_Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN VARCHAR2     := null,
   numFmtId_      IN PLS_INTEGER  := null,
   fontId_        IN PLS_INTEGER  := null,
   fillId_        IN PLS_INTEGER  := null,
   borderId_      IN PLS_INTEGER  := null,
   alignment_     IN tp_alignment := null,
   sheet_         IN PLS_INTEGER := null );

PROCEDURE Mergecells (
   tl_col_ IN PLS_INTEGER, -- top left
   tl_row_ IN PLS_INTEGER,
   br_col_ IN PLS_INTEGER, -- bottom right
   br_row_ IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null );

PROCEDURE List_Validation (
   p_sqref_col    IN PLS_INTEGER,
   p_sqref_row    IN PLS_INTEGER,
   p_tl_col       IN PLS_INTEGER, -- top left
   p_tl_row       IN PLS_INTEGER,
   p_br_col       IN PLS_INTEGER, -- bottom right
   p_br_row       IN PLS_INTEGER,
   p_style        IN VARCHAR2 := 'stop', -- stop, warning, information
   p_title        IN VARCHAR2 := null,
   p_prompt       IN VARCHAR  := null,
   p_show_error   IN BOOLEAN  := false,
   p_error_title  IN VARCHAR2 := null,
   p_error_txt    IN VARCHAR2 := null,
   sheet_        IN PLS_INTEGER := null );

PROCEDURE List_Validation (
   p_sqref_col    IN PLS_INTEGER,
   p_sqref_row    IN PLS_INTEGER,
   p_defined_name IN VARCHAR2,
   p_style        IN VARCHAR2 := 'stop', -- stop, warning, information
   p_title        IN VARCHAR2 := null,
   p_prompt       IN VARCHAR  := null,
   p_show_error   IN BOOLEAN  := false,
   p_error_title  IN VARCHAR2 := null,
   p_error_txt    IN VARCHAR2 := null,
   sheet_        IN PLS_INTEGER := null );

PROCEDURE Defined_Name (
   tl_col_     IN PLS_INTEGER, -- top left
   tl_row_     IN PLS_INTEGER,
   br_col_     IN PLS_INTEGER, -- bottom right
   br_row_     IN PLS_INTEGER,
   name_       IN VARCHAR2,
   sheet_      IN PLS_INTEGER := null,
   localsheet_ IN PLS_INTEGER := null );

PROCEDURE Set_Column_Width (
   col_   IN PLS_INTEGER,
   width_ IN NUMBER,
   sheet_ IN PLS_INTEGER := null );

PROCEDURE Set_Column (
   col_       IN PLS_INTEGER,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null );

PROCEDURE Set_Row (
   row_       IN PLS_INTEGER,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null,
   height_    IN NUMBER       := null );

PROCEDURE Freeze_Rows (
   nr_rows_  IN PLS_INTEGER := 1,
   sheet_    IN PLS_INTEGER := null );

PROCEDURE Freeze_Cols (
   nr_cols_ IN PLS_INTEGER := 1,
   sheet_   IN PLS_INTEGER := null );

PROCEDURE Freeze_Pane (
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER,
   sheet_ IN PLS_INTEGER := null );

PROCEDURE Set_Autofilter (
   col_start_ IN PLS_INTEGER := null,
   col_end_   IN PLS_INTEGER := null,
   row_start_ IN PLS_INTEGER := null,
   row_end_   IN PLS_INTEGER := null,
   sheet_     IN PLS_INTEGER := null );

PROCEDURE Set_Tabcolor (
   tabcolor_ VARCHAR2, -- hex Alpha-rgb value
   sheet_    PLS_INTEGER := null );

FUNCTION Finish RETURN BLOB;

PROCEDURE Save (
   directory_ VARCHAR2,
   filename_  VARCHAR2 );

PROCEDURE Save (
   xl_blob_   IN BLOB,
   directory_ IN VARCHAR2,
   filename_  IN VARCHAR2 );

PROCEDURE Query2Sheet (
   col_count_   IN OUT PLS_INTEGER,
   row_count_   IN OUT PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() );

PROCEDURE Query2Sheet (
   col_count_   IN OUT PLS_INTEGER,
   row_count_   IN OUT PLS_INTEGER,
   sql_         IN VARCHAR2,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() );

PROCEDURE Query2Sheet ( -- using REFCURSOR
   col_count_   IN OUT PLS_INTEGER,
   row_count_   IN OUT PLS_INTEGER,
   rc_          IN OUT SYS_REFCURSOR,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() );

PROCEDURE Query2SheetAndAutofilter ( -- with Binds
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() );

PROCEDURE Query2SheetAndAutofilter ( -- no Binds
   sql_         IN VARCHAR2,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() );

PROCEDURE SetUseXf (
   p_val BOOLEAN := true );

------------------------------------------------------------------------------
-- Special Page Generators
--
PROCEDURE Set_Param (
   params_ IN OUT params_arr,
   ix_     IN NUMBER,
   val_    IN VARCHAR2,
   extra_  IN VARCHAR2 := '' );

PROCEDURE Bind_Value (
   binds_   IN OUT bind_arr,
   bind_id_ IN VARCHAR2,
   val_     IN VARCHAR2 );
PROCEDURE Bind_Value (
   binds_   IN OUT bind_arr,
   bind_id_ IN VARCHAR2,
   val_     IN NUMBER );
PROCEDURE Bind_Value (
   binds_   IN OUT bind_arr,
   bind_id_ IN VARCHAR2,
   val_     IN DATE );

PROCEDURE Create_Params_Sheet (
   report_name_ IN VARCHAR2,
   params_      IN params_arr,
   show_user_   IN BOOLEAN     := true,
   sheet_       IN PLS_INTEGER := null );


END AS_XLSX;
/
CREATE OR REPLACE PACKAGE BODY AS_XLSX IS

VERSION_ CONSTANT VARCHAR2(20) := 'as_xlsx20';

LOCAL_FILE_HEADER_        CONSTANT RAW(4) := hextoraw('504B0304'); -- Local file header signature
END_OF_CENTRAL_DIRECTORY_ CONSTANT RAW(4) := hextoraw('504B0506'); -- End of central directory signature

TYPE tp_XF_fmt IS RECORD (
   numFmtId  PLS_INTEGER,
   fontId    PLS_INTEGER,
   fillId    PLS_INTEGER,
   borderId  PLS_INTEGER,
   alignment tp_alignment,
   height    NUMBER
);
TYPE tp_col_fmts is table of tp_XF_fmt index by PLS_INTEGER;
TYPE tp_row_fmts is table of tp_XF_fmt index by PLS_INTEGER;
TYPE tp_widths is table of NUMBER index by PLS_INTEGER;
TYPE tp_cell IS RECORD (
   value       NUMBER,
   style       VARCHAR2(50),
   formula_idx PLS_INTEGER
);
TYPE tp_cells is table of tp_cell index by PLS_INTEGER;
TYPE tp_rows is table of tp_cells index by PLS_INTEGER;
TYPE tp_autofilter is record (
   column_start PLS_INTEGER,
   column_end   PLS_INTEGER,
   row_start    PLS_INTEGER,
   row_end PLS_INTEGER
);
TYPE tp_autofilters is table of tp_autofilter index by PLS_INTEGER;
TYPE tp_hyperlink is record (
   cell VARCHAR2(10),
   url  VARCHAR2(1000)
);
TYPE tp_hyperlinks is table of tp_hyperlink index by PLS_INTEGER;
SUBTYPE tp_author is VARCHAR2(32767 char);
TYPE tp_authors is table of PLS_INTEGER index by tp_author;
authors tp_authors;
TYPE tp_formulas is table of VARCHAR2(32767) index by PLS_INTEGER;
TYPE tp_comment is record (
   text   VARCHAR2(32767 char),
   author tp_author,
   row    PLS_INTEGER,
   column PLS_INTEGER,
   width  PLS_INTEGER,
   height PLS_INTEGER
);
TYPE tp_comments   is table of tp_comment index by PLS_INTEGER;
TYPE tp_mergecells is table of VARCHAR2(21) index by PLS_INTEGER;
TYPE tp_validation IS RECORD (
   type             VARCHAR2(10),
   errorstyle       VARCHAR2(32),
   showinputmessage BOOLEAN,
   prompt           VARCHAR2(32767 CHAR),
   title            VARCHAR2(32767 CHAR),
   error_title      VARCHAR2(32767 CHAR),
   error_txt        VARCHAR2(32767 CHAR),
   showerrormessage BOOLEAN,
   formula1         VARCHAR2(32767 CHAR),
   formula2         VARCHAR2(32767 CHAR),
   allowBlank       BOOLEAN,
   sqref            VARCHAR2(32767 CHAR)
);
TYPE tp_validations IS TABLE OF tp_validation INDEX BY PLS_INTEGER;
TYPE tp_sheet IS RECORD (
   rows        tp_rows,
   widths      tp_widths,
   name        VARCHAR2(100),
   freeze_rows PLS_INTEGER,
   freeze_cols PLS_INTEGER,
   autofilters tp_autofilters,
   hyperlinks  tp_hyperlinks,
   col_fmts    tp_col_fmts,
   row_fmts    tp_row_fmts,
   comments    tp_comments,
   mergecells  tp_mergecells,
   validations tp_validations,
   tabcolor    VARCHAR2(8),
   fontid      PLS_INTEGER
);
TYPE tp_sheets is table of tp_sheet index by PLS_INTEGER;
TYPE tp_numFmt IS RECORD (
   numFmtId   PLS_INTEGER,
   formatCode VARCHAR2(100)
);
TYPE tp_numFmts is table of tp_numFmt index by PLS_INTEGER;
TYPE tp_fill is record (
   patternType VARCHAR2(30),
   fgRGB VARCHAR2(8),
   bgRGB VARCHAR2(8)
);
TYPE tp_fills is table of tp_fill index by PLS_INTEGER;
TYPE tp_cellXfs is table of tp_xf_fmt index by PLS_INTEGER;
TYPE tp_font is record (
   name      VARCHAR2(100),
   family    PLS_INTEGER,
   fontsize  NUMBER,
   theme     PLS_INTEGER,
   RGB       VARCHAR2(8),
   underline BOOLEAN,
   italic    BOOLEAN,
   bold BOOLEAN
);
TYPE tp_fonts is table of tp_font index by PLS_INTEGER;
TYPE tp_border is record (
   top    VARCHAR2(17),
   bottom VARCHAR2(17),
   left   VARCHAR2(17),
   right VARCHAR2(17)
);
TYPE tp_borders is table of tp_border index by PLS_INTEGER;
TYPE tp_numFmtIndexes is table of PLS_INTEGER index by PLS_INTEGER;
TYPE tp_strings is table of PLS_INTEGER index by VARCHAR2(32767 char);
TYPE tp_str_ind is table of VARCHAR2(32767 char) index by PLS_INTEGER;
TYPE tp_defined_name is record (
   name VARCHAR2(32767 char),
   ref VARCHAR2(32767 char),
   sheet PLS_INTEGER
);
TYPE tp_defined_names is table of tp_defined_name index by PLS_INTEGER;
TYPE tp_book IS RECORD (
   sheets        tp_sheets,
   strings       tp_strings,
   str_ind       tp_str_ind,
   str_cnt       PLS_INTEGER := 0,
   fonts         tp_fonts,
   fills         tp_fills,
   borders       tp_borders,
   numFmts       tp_numFmts,
   cellXfs       tp_cellXfs,
   numFmtIndexes tp_numFmtIndexes,
   defined_names tp_defined_names,
   formulas      tp_formulas,
   fontid        PLS_INTEGER
);

workbook              tp_book;
g_useXf_              BOOLEAN := true;
g_addtxt2utf8blob_tmp VARCHAR2(32767);

PROCEDURE addtxt2utf8blob_init (
   blob_ IN OUT NOCOPY BLOB )
IS BEGIN
   g_addtxt2utf8blob_tmp := null;
   dbms_lob.createtemporary (blob_, true);
END addtxt2utf8blob_init;

PROCEDURE Addtxt2utf8blob_Finish (
   blob_ IN OUT NOCOPY BLOB )
IS
   raw_ RAW(32767);
BEGIN
   raw_ := utl_i18n.string_to_raw (g_addtxt2utf8blob_tmp, 'AL32UTF8');
   dbms_lob.writeappend (blob_, utl_raw.length(raw_), raw_);
EXCEPTION
   WHEN value_error THEN
      raw_ := utl_i18n.string_to_raw(substr(g_addtxt2utf8blob_tmp,1,16381), 'AL32UTF8');
      dbms_lob.writeappend (blob_, utl_raw.length(raw_), raw_);
      raw_ := utl_i18n.string_to_raw(substr(g_addtxt2utf8blob_tmp,16382), 'AL32UTF8');
      dbms_lob.writeappend (blob_, utl_raw.length(raw_), raw_);
END Addtxt2utf8blob_Finish;

PROCEDURE addtxt2utf8blob( p_txt VARCHAR2, p_blob in out nocopy blob )
IS BEGIN
   g_addtxt2utf8blob_tmp := g_addtxt2utf8blob_tmp || p_txt;
EXCEPTION
   WHEN value_error THEN
      addtxt2utf8blob_finish(p_blob);
      g_addtxt2utf8blob_tmp := p_txt;
END addtxt2utf8blob;

PROCEDURE Blob2File (
   blob_      BLOB,
   directory_ VARCHAR2 := 'MY_DIR',
   filename_  VARCHAR2 := 'my.xlsx' )
IS
   fh_  utl_file.file_type;
   len_ PLS_INTEGER := 32767;
BEGIN
   fh_ := Utl_File.fopen (directory_, filename_, 'wb');
   FOR i_ IN 0 .. trunc((dbms_lob.getlength(blob_)-1)/len_) LOOP
      Utl_File.Put_Raw (fh_, dbms_lob.substr(blob_, len_, i_*len_+1));
   END LOOP;
   Utl_File.fclose (fh_);
END Blob2File;

FUNCTION Raw2Num (
   raw_ RAW,
   len_ INTEGER,
   pos_ INTEGER ) RETURN NUMBER
IS BEGIN
   RETURN utl_raw.cast_to_binary_integer(
      utl_raw.substr(raw_, pos_, len_),
      utl_raw.little_endian
   );
END Raw2Num;

FUNCTION Little_Endian (
   big_   NUMBER,
   bytes_ PLS_INTEGER := 4 ) RETURN RAW
IS BEGIN
   RETURN utl_raw.substr (
      utl_raw.cast_from_binary_integer (big_, utl_raw.little_endian),
      1, bytes_
   );
END Little_Endian;

FUNCTION Blob2Num (
   blob_ BLOB,
   len_  INTEGER,
   pos_  INTEGER ) RETURN NUMBER
IS BEGIN
   RETURN utl_raw.cast_to_binary_integer (
      dbms_lob.substr(blob_, len_, pos_),
      utl_raw.little_endian
   );
END Blob2Num;

PROCEDURE Add1File (
   zipped_blob_ IN OUT BLOB,
   name_        IN VARCHAR2,
   content_     IN BLOB )
IS
   now_        DATE := sysdate;
   blob_       BLOB;
   len_        INTEGER;
   clen_       INTEGER;
   crc32_      RAW(4) := hextoraw( '00000000' );
   compressed_ BOOLEAN := false;
   name_raw_   RAW(32767);
BEGIN
   len_ := nvl(Dbms_Lob.GetLength( content_ ), 0 );
   IF len_ > 0 THEN
      blob_       := Utl_Compress.Lz_Compress (content_);
      clen_       := Dbms_Lob.GetLength (blob_)-18;
      compressed_ := clen_ < len_;
      crc32_      := Dbms_Lob.Substr (blob_, 4, clen_+11);
   END IF;
   IF not compressed_ THEN
      clen_ := len_;
      blob_ := content_;
   END IF;
   IF zipped_blob_ IS null THEN
      dbms_lob.createtemporary( zipped_blob_, true );
   END IF;
   name_raw_ := Utl_i18n.String_To_Raw (name_, 'AL32UTF8');
   Dbms_Lob.Append (
      zipped_blob_,
      Utl_Raw.Concat(
         LOCAL_FILE_HEADER_, -- Local file header signature
         hextoraw('1400'),   -- version 2.0
         CASE WHEN name_raw_ = Utl_i18n.String_To_Raw (name_, 'US8PC437')
            THEN hextoraw('0000') -- no General purpose bits
            ELSE hextoraw('0008') -- set Language encoding flag (EFS)
         END, CASE WHEN compressed_
            THEN hextoraw('0800') -- deflate
            ELSE hextoraw('0000') -- stored
         END,
         Little_Endian (to_number(to_char (now_, 'ss'))/2
                        + to_number(to_char (now_, 'mi'))*32
                        + to_number(to_char (now_, 'hh24'))*2048, 2), -- File last modification time
         Little_Endian (to_number(to_char(now_,'dd'))
                        + to_number(to_char(now_,'mm'))*32
                        + (to_number(to_char(now_,'yyyy'))-1980)*512, 2), -- File last modification date
         crc32_,               -- CRC-32
         Little_Endian(clen_), -- compressed size
         Little_Endian(len_),  -- uncompressed size
         Little_Endian (Utl_Raw.Length(name_raw_), 2), -- File name length
         hextoraw('0000'),     -- Extra field length
         name_raw_             -- File name
      )
   );
   IF compressed_ THEN
      dbms_lob.copy( zipped_blob_, blob_, clen_, dbms_lob.getlength( zipped_blob_ ) + 1, 11 ); -- compressed content
   ELSIF clen_ > 0 THEN
      dbms_lob.copy( zipped_blob_, blob_, clen_, dbms_lob.getlength( zipped_blob_ ) + 1, 1 ); --  content
   END IF;
   IF dbms_lob.istemporary(blob_) = 1 THEN
      Dbms_Lob.FreeTemporary (blob_);
   END IF;
END Add1File;

PROCEDURE Finish_Zip (
   zipped_blob_ IN OUT BLOB )
IS
   t_cnt             PLS_INTEGER := 0;
   t_offs            INTEGER;
   t_offs_dir_header INTEGER;
   t_offs_end_header INTEGER;
   t_comment         RAW(200) := Utl_Raw.Cast_To_Raw (
      'Implementation by Anton Scheffer, ' || VERSION_
   );
BEGIN
   t_offs_dir_header := dbms_lob.getlength (zipped_blob_);
   t_offs := 1;
   WHILE Dbms_Lob.Substr(zipped_blob_, utl_raw.length(LOCAL_FILE_HEADER_), t_offs) = LOCAL_FILE_HEADER_ LOOP
      t_cnt := t_cnt + 1;
      Dbms_Lob.Append (
         zipped_blob_,
         Utl_Raw.Concat (
            hextoraw('504B0102'),      -- Central directory file header signature
            hextoraw('1400'),          -- version 2.0
            dbms_lob.substr(zipped_blob_, 26, t_offs+4),
            hextoraw('0000'),          -- File comment length
            hextoraw('0000'),          -- Disk number where file starts
            hextoraw('0000'),          -- Internal file attributes => 0000=binary-file; 0100(ascii)=text-file
            CASE
               WHEN Dbms_Lob.Substr (
                  zipped_blob_, 1, t_offs+30+blob2num(zipped_blob_,2,t_offs+26)-1
               ) IN (hextoraw('2F'), hextoraw('5C'))
               THEN
                  hextoraw('10000000') -- a directory/folder
               ELSE
                  hextoraw('2000B681') -- a file
            END,                       -- External file attributes
            little_endian(t_offs-1),   -- Relative offset of local file header
            dbms_lob.substr(zipped_blob_, blob2num(zipped_blob_,2,t_offs+26),t_offs+30) -- File name
         )
      );
      t_offs := t_offs + 30 +
         blob2num(zipped_blob_, 4, t_offs+18 ) + -- compressed size
         blob2num(zipped_blob_, 2, t_offs+26 ) + -- File name length
         blob2num(zipped_blob_, 2, t_offs+28 );  -- Extra field length
   END LOOP;
   t_offs_end_header := dbms_lob.getlength(zipped_blob_);
   Dbms_Lob.Append (
       zipped_blob_,
       Utl_Raw.Concat (
          END_OF_CENTRAL_DIRECTORY_,                          -- End of central directory signature
          hextoraw('0000'),                                   -- Number of this disk
          hextoraw('0000'),                                   -- Disk where central directory starts
          little_endian(t_cnt,2),                             -- Number of central directory records on this disk
          little_endian(t_cnt,2),                             -- Total number of central directory records
          little_endian(t_offs_end_header-t_offs_dir_header), -- Size of central directory
          little_endian(t_offs_dir_header),                   -- Offset of start of central directory, relative to start of archive
          little_endian(nvl(Utl_Raw.Length(t_comment),0),2),  -- ZIP file comment length
          t_comment
       )
    );
END Finish_Zip;

FUNCTION Alfan_Col (
   p_col PLS_INTEGER ) RETURN VARCHAR2
IS BEGIN
   RETURN CASE
      WHEN p_col > 702 THEN chr(64+trunc((p_col-27)/676)) || chr(65+mod(trunc((p_col-1)/26)-1, 26)) || chr(65+mod(p_col-1, 26))
      WHEN p_col > 26  THEN chr(64+trunc((p_col-1)/26)) || chr(65+mod(p_col-1, 26))
      ELSE chr(64+p_col)
   END;
END Alfan_Col;

FUNCTION Col_Alfan(
   p_col VARCHAR2 ) RETURN PLS_INTEGER
IS BEGIN
   RETURN ascii(substr(p_col,-1)) - 64
      + nvl((ascii(substr(p_col,-2,1))-64) * 26, 0)
      + nvl((ascii(substr(p_col,-3,1))-64) * 676, 0);
END Col_Alfan;

PROCEDURE Clear_Workbook
IS
   s_      PLS_INTEGER := workbook.sheets.first;
   row_ix_ PLS_INTEGER;
BEGIN
   WHILE s_ IS NOT null LOOP
      row_ix_ := workbook.sheets(s_).rows.first();
      WHILE row_ix_ IS NOT null LOOP
         workbook.sheets(s_).rows(row_ix_).delete();
         row_ix_ := workbook.sheets(s_).rows.next(row_ix_);
      END LOOP;
      workbook.sheets(s_).rows.delete();
      workbook.sheets(s_).widths.delete();
      workbook.sheets(s_).autofilters.delete();
      workbook.sheets(s_).hyperlinks.delete();
      workbook.sheets(s_).col_fmts.delete();
      workbook.sheets(s_).row_fmts.delete();
      workbook.sheets(s_).comments.delete();
      workbook.sheets(s_).mergecells.delete();
      workbook.sheets(s_).validations.delete();
      s_ := workbook.sheets.next(s_);
   END LOOP;
   workbook.strings.delete();
   workbook.str_ind.delete();
   workbook.fonts.delete();
   workbook.fills.delete();
   workbook.borders.delete();
   workbook.numFmts.delete();
   workbook.cellXfs.delete();
   workbook.defined_names.delete();
   workbook.formulas.delete();
   workbook := null;
END Clear_Workbook;

PROCEDURE Set_Tabcolor (
   tabcolor_ VARCHAR2,
   sheet_    PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).tabcolor := substr(tabcolor_, 1, 8);
END Set_Tabcolor;


FUNCTION New_Sheet (
   sheetname_ VARCHAR2 := null,
   tab_color_ VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   s_ PLS_INTEGER := workbook.sheets.count() + 1;
BEGIN
   workbook.sheets(s_).name := nvl(dbms_xmlgen.convert(translate(sheetname_, 'a/\[]*:?', 'a')), 'Sheet'||s_);
   IF workbook.strings.count() = 0 THEN
      workbook.str_cnt := 0;
   END IF;
   IF workbook.fonts.count() = 0 THEN
      workbook.fontid := Get_Font('Calibri');
   END IF;
   IF workbook.fills.count() = 0 THEN
      Get_Fill('none');
      Get_Fill('gray125');
   END IF;
   IF workbook.borders.count() = 0 THEN
      Get_Border ('', '', '', '');
   END IF;
   Set_TabColor(tab_color_, s_);
   workbook.sheets(s_).fontid := workbook.fontid;
   RETURN s_;
END New_Sheet;

PROCEDURE New_Sheet (
   sheetname_ VARCHAR2 := null,
   tab_color_ VARCHAR2 := null )
IS
   s_ PLS_INTEGER := New_Sheet (sheetname_, tab_color_); --ignore
BEGIN
   null;
END New_Sheet;

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 )
IS BEGIN
   workbook.sheets(sheet_).name := nvl(dbms_xmlgen.convert(translate(name_, 'a/\[]*:?', 'a')), 'Sheet'||sheet_);
END Set_Sheet_Name;

PROCEDURE Set_Col_Width (
   sheet_  PLS_INTEGER,
   col_    PLS_INTEGER,
   format_ VARCHAR2 )
IS
   width_  NUMBER;
   nr_chr_ PLS_INTEGER;
BEGIN
   IF format_ IS null THEN
      RETURN;
   END IF;
   IF instr(format_, ';') > 0 THEN
      nr_chr_ := length(translate(substr(format_, 1, instr(format_,';')-1), 'a\"', 'a'));
   ELSE
      nr_chr_ := length(translate(format_, 'a\"', 'a'));
   END IF;
   width_ := trunc((nr_chr_*7+5)/7*256)/256; -- assume default 11 point Calibri
   IF workbook.sheets(sheet_).widths.exists(col_) THEN
      workbook.sheets(sheet_).widths(col_) := greatest(
         workbook.sheets(sheet_).widths(col_), width_
      );
   ELSE
      workbook.sheets(sheet_).widths(col_) := greatest(width_, 8.43);
   END IF;
END Set_Col_Width;


FUNCTION OraFmt2Excel (
   p_format VARCHAR2 := null ) RETURN VARCHAR2
IS
   t_format VARCHAR2(1000) := substr (p_format, 1, 1000);
BEGIN
   t_format := replace(replace(t_format,'hh24','hh'),'hh12','hh');
   t_format := replace( t_format, 'mi', 'mm' );
   t_format := replace( replace( replace( t_format, 'AM', '~~' ), 'PM', '~~' ), '~~', 'AM/PM' );
   t_format := replace( replace( replace( t_format, 'am', '~~' ), 'pm', '~~' ), '~~', 'AM/PM' );
   t_format := replace( replace( t_format, 'day', 'DAY' ), 'DAY', 'dddd' );
   t_format := replace( replace( t_format, 'dy', 'DY' ), 'DAY', 'ddd' );
   t_format := replace( replace( t_format, 'RR', 'RR' ), 'RR', 'YY' );
   t_format := replace( replace( t_format, 'month', 'MONTH' ), 'MONTH', 'mmmm' );
   t_format := replace( replace( t_format, 'mon', 'MON' ), 'MON', 'mmm' );
   t_format := replace( t_format, '9', '#' );
   t_format := replace( t_format, 'D', '.' );
   t_format := replace( t_format, 'G', ',' );
   RETURN t_format;
END OraFmt2Excel;

FUNCTION Get_NumFmt (
   format_ VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   cnt_      PLS_INTEGER;
   numFmtId_ PLS_INTEGER;
BEGIN
   IF format_ is null THEN
      RETURN 0;
   END IF;
   cnt_ := workbook.numFmts.count();
   FOR i_ in 1 .. cnt_ LOOP
      IF workbook.numFmts(i_).formatCode = format_ THEN
         numFmtId_ := workbook.numFmts(i_).numFmtId;
         EXIT;
      END IF;
   END LOOP;
   IF numFmtId_ is null THEN
      numFmtId_ := CASE WHEN cnt_ = 0 THEN 164 ELSE workbook.numFmts(cnt_).numFmtId + 1 END;
      cnt_ := cnt_ + 1;
      workbook.numFmts(cnt_).numFmtId   := numFmtId_;
      workbook.numFmts(cnt_).formatCode := format_;
      workbook.numFmtIndexes(numFmtId_) := cnt_;
   END IF;
   RETURN numFmtId_;
END Get_NumFmt;

PROCEDURE Set_Font (
   name_      VARCHAR2    := 'Calibri',
   sheet_     PLS_INTEGER := null,
   family_    PLS_INTEGER := 2,
   fontsize_  NUMBER      := 11,
   theme_     PLS_INTEGER := 1,
   underline_ BOOLEAN     := false,
   italic_    BOOLEAN     := false,
   bold_      BOOLEAN     := false,
   rgb_       VARCHAR2    := null ) -- this is a hex ALPHA Red Green Blue value
IS
   ix_ PLS_INTEGER := Get_Font (name_, family_, fontsize_, theme_, underline_, italic_, bold_, rgb_);
BEGIN
   IF sheet_ IS null THEN
      workbook.fontid := ix_;
   ELSE
      workbook.sheets(sheet_).fontid := ix_;
   END IF;
END Set_Font;

FUNCTION Get_Font (
   name_      VARCHAR2    := 'Calibri',
   family_    PLS_INTEGER := 2,
   fontsize_  NUMBER      := 11,
   theme_     PLS_INTEGER := 1,
   underline_ BOOLEAN     := false,
   italic_    BOOLEAN     := false,
   bold_      BOOLEAN     := false,
   rgb_       VARCHAR2    := null ) RETURN PLS_INTEGER
IS
   ix_ PLS_INTEGER;
BEGIN
   IF workbook.fonts.count() > 0 THEN
      FOR f_ IN 0 .. workbook.fonts.count() - 1 LOOP
         IF (     workbook.fonts(f_).name      = name_
              AND workbook.fonts(f_).family    = family_
              AND workbook.fonts(f_).fontsize  = fontsize_
              AND workbook.fonts(f_).theme     = theme_
              AND workbook.fonts(f_).underline = underline_
              AND workbook.fonts(f_).italic    = italic_
              AND workbook.fonts(f_).bold      = bold_
              AND (     workbook.fonts(f_).rgb = rgb_
                    OR (workbook.fonts(f_).rgb IS null AND rgb_ IS null)
              )
         ) THEN
            RETURN f_;
         END IF;
      END LOOP;
   END IF;
   ix_ := workbook.fonts.count();
   workbook.fonts(ix_).name      := name_;
   workbook.fonts(ix_).family    := family_;
   workbook.fonts(ix_).fontsize  := fontsize_;
   workbook.fonts(ix_).theme     := theme_;
   workbook.fonts(ix_).underline := underline_;
   workbook.fonts(ix_).italic    := italic_;
   workbook.fonts(ix_).bold      := bold_;
   workbook.fonts(ix_).rgb       := rgb_;
   RETURN ix_;
END Get_Font;


FUNCTION Get_Fill (
   patternType_ VARCHAR2,
   fgRGB_       VARCHAR2 := null,
   bgRGB_       VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   ix_ PLS_INTEGER;
BEGIN
   IF workbook.fills.count() > 0 THEN
      FOR f_ IN 0 .. workbook.fills.count() - 1 LOOP
         IF (   workbook.fills(f_).patternType = patternType_
            AND nvl(workbook.fills(f_).fgRGB, 'x') = nvl(upper(fgRGB_), 'x')
            AND nvl(workbook.fills(f_).bgRGB, 'x') = nvl(upper(bgRGB_), 'x')
         ) THEN
            RETURN f_;
         END IF;
      END LOOP;
   END IF;
   ix_ := workbook.fills.count();
   workbook.fills(ix_).patternType := patternType_;
   workbook.fills(ix_).fgRGB       := upper(fgRGB_);
   workbook.fills(ix_).bgRGB       := upper(bgRGB_);
   RETURN ix_;
END Get_Fill;

PROCEDURE Get_Fill (
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,
   bgRGB_       IN VARCHAR2 := null )
IS
   ix_ PLS_INTEGER := Get_Fill (patternType_, fgRGB_, bgRGB_); --ignore
BEGIN
   null;
END Get_Fill;


FUNCTION Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' ) RETURN PLS_INTEGER
IS
   ix_ PLS_INTEGER;
BEGIN
   IF workbook.borders.count() > 0 THEN
      FOR b_ IN 0 .. workbook.borders.count() - 1 LOOP
         IF (   nvl(workbook.borders(b_).top,    'x') = nvl(top_, 'x')
            AND nvl(workbook.borders(b_).bottom, 'x') = nvl(bottom_, 'x')
            AND nvl(workbook.borders(b_).left,   'x') = nvl(left_, 'x')
            AND nvl(workbook.borders(b_).right,  'x') = nvl(right_, 'x')
         ) THEN
            RETURN b_;
         END IF;
      END LOOP;
   END IF;
   ix_ := workbook.borders.count();
   workbook.borders(ix_).top    := top_;
   workbook.borders(ix_).bottom := bottom_;
   workbook.borders(ix_).left   := left_;
   workbook.borders(ix_).right  := right_;
   RETURN ix_;
END Get_Border;

PROCEDURE Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' )
IS
   ix_ NUMBER := Get_Border (top_, bottom_, left_, right_);
BEGIN
   SELECT 1 INTO ix_ FROM dual; -- avoid compiler warning
END Get_Border;

FUNCTION Get_Alignment (
   vertical_    VARCHAR2 := null,
   horizontal_  VARCHAR2 := null,
   wrapText_    BOOLEAN := null ) RETURN tp_alignment
IS
   rv_ tp_alignment;
BEGIN
   rv_.vertical := vertical_;
   rv_.horizontal := horizontal_;
   rv_.wrapText := wrapText_;
   RETURN rv_;
END Get_Alignment;

FUNCTION Get_XfId (
   sheet_     PLS_INTEGER,
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   numFmtId_  PLS_INTEGER := null,
   fontId_    PLS_INTEGER := null,
   fillId_    PLS_INTEGER := null,
   borderId_  PLS_INTEGER := null,
   alignment_ tp_alignment := null ) RETURN VARCHAR2
IS
   cnt_    PLS_INTEGER;
   XfId_   PLS_INTEGER;
   XF_     tp_XF_fmt;
   col_XF_ tp_XF_fmt;
   row_XF_ tp_XF_fmt;
BEGIN
   IF not g_useXf_ THEN
      RETURN '';
   END IF;
   IF workbook.sheets(sheet_).col_fmts.exists(col_) THEN
      col_XF_ := workbook.sheets(sheet_).col_fmts(col_);
   END IF;
   IF workbook.sheets(sheet_).row_fmts.exists(row_) THEN
      row_XF_ := workbook.sheets(sheet_).row_fmts(row_);
   END IF;
   XF_.numFmtId := coalesce (numFmtId_, col_XF_.numFmtId, row_XF_.numFmtId, workbook.sheets(sheet_).fontid, workbook.fontid);
   XF_.fontId   := coalesce (fontId_, col_XF_.fontId, row_XF_.fontId, 0);
   XF_.fillId   := coalesce (fillId_, col_XF_.fillId, row_XF_.fillId, 0);
   XF_.borderId := coalesce (borderId_, col_XF_.borderId, row_XF_.borderId, 0);
   XF_.alignment := Get_Alignment (
      coalesce (alignment_.vertical, col_XF_.alignment.vertical, row_XF_.alignment.vertical),
      coalesce (alignment_.horizontal, col_XF_.alignment.horizontal, row_XF_.alignment.horizontal),
      coalesce (alignment_.wrapText, col_XF_.alignment.wrapText, row_XF_.alignment.wrapText)
   );
   IF XF_.numFmtId + XF_.fontId + XF_.fillId + XF_.borderId = 0
      AND XF_.alignment.vertical IS null
      AND XF_.alignment.horizontal IS null
      AND not nvl(XF_.alignment.wrapText, false)
   THEN
      RETURN '';
   END IF;
   IF XF_.numFmtId > 0 THEN
      Set_Col_Width (sheet_, col_, workbook.numFmts(workbook.numFmtIndexes(XF_.numFmtId)).formatCode);
   END IF;
   cnt_ := workbook.cellXfs.count();
   FOR i IN 1 .. cnt_ LOOP
      IF (   workbook.cellXfs(i).numFmtId = XF_.numFmtId
         and workbook.cellXfs(i).fontId = XF_.fontId
         and workbook.cellXfs(i).fillId = XF_.fillId
         and workbook.cellXfs(i).borderId = XF_.borderId
         and nvl(workbook.cellXfs(i).alignment.vertical, 'x') = nvl (XF_.alignment.vertical, 'x')
         and nvl(workbook.cellXfs(i).alignment.horizontal, 'x') = nvl (XF_.alignment.horizontal, 'x')
         and nvl(workbook.cellXfs(i).alignment.wrapText, false) = nvl (XF_.alignment.wrapText, false)
      ) THEN
         XfId_ := i;
         EXIT;
      END IF;
   END LOOP;
   IF XfId_ IS null THEN
      cnt_ := cnt_ + 1;
      XfId_ := cnt_;
      workbook.cellXfs(cnt_) := XF_;
   END IF;
   RETURN 's="' || XfId_ || '"';
END Get_XfId;

FUNCTION Extract_Id_From_Style (
   style_ IN VARCHAR2 ) RETURN PLS_INTEGER
IS BEGIN
   RETURN CASE
      WHEN style_ IS null OR style_ = 't="s" ' THEN null
      ELSE to_number(regexp_replace (style_, 't="s" s="(\d+)"', '\1'))
   END;
END Extract_Id_From_Style;

PROCEDURE Cell (
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN NUMBER,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).rows(row_)(col_).value := value_;
   workbook.sheets(sh_).rows(row_)(col_).style := null;
   workbook.sheets(sh_).rows(row_)(col_).style := get_XfId (
      sh_, col_, row_, numFmtId_, fontId_, fillId_, borderId_, alignment_
   );
END Cell;

PROCEDURE CellP (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     NUMBER,
   numFmtId_  VARCHAR2    := null,
   fontId_    VARCHAR2    := null,
   fillId_    VARCHAR2    := null,
   borderId_  VARCHAR2    := null,
   alignment_ VARCHAR2    := null,
   sheet_     PLS_INTEGER := null )
IS BEGIN
   Cell (
      col_, row_, value_,
      CASE WHEN numFmtId_  IS NOT null THEN numFmt_(numFmtId_) END,
      CASE WHEN fontId_    IS NOT null THEN fonts_(fontId_) END,
      CASE WHEN fillId_    IS NOT null THEN fills_(fillId_) END,
      CASE WHEN borderId_  IS NOT null THEN bdrs_(borderId_) END,
      CASE WHEN alignment_ IS NOT null THEN align_(alignment_) END,
      sheet_
   );
END CellP;

FUNCTION Add_String (
   string_ VARCHAR2 ) RETURN PLS_INTEGER
IS
   cnt_ PLS_INTEGER;
BEGIN
   IF workbook.strings.exists(nvl(string_, '')) THEN
      cnt_ := workbook.strings(nvl(string_, ''));
   ELSE
      cnt_ := workbook.strings.count();
      workbook.str_ind(cnt_) := string_;
      workbook.strings(nvl(string_, '')) := cnt_;
   END IF;
   workbook.str_cnt := workbook.str_cnt + 1;
   RETURN cnt_;
END Add_String;

PROCEDURE Cell (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     VARCHAR2,
   numFmtId_  PLS_INTEGER  := null,
   fontId_    PLS_INTEGER  := null,
   fillId_    PLS_INTEGER  := null,
   borderId_  PLS_INTEGER  := null,
   alignment_ tp_alignment := null,
   sheet_     PLS_INTEGER  := null )
IS
   sh_    PLS_INTEGER  := nvl(sheet_, workbook.sheets.count());
   align_ tp_alignment := alignment_;
BEGIN
   workbook.sheets(sh_).rows(row_)(col_).value := Add_String(value_);
   IF align_.wrapText IS null AND instr(value_, chr(13)) > 0 THEN
      align_.wrapText := true;
   END IF;
   workbook.sheets(sh_).rows(row_)(col_).style := 't="s" ' || get_XfId (
      sh_, col_, row_, numFmtId_, fontId_, fillId_, borderId_, align_
   );
END Cell;

PROCEDURE CellP (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     VARCHAR2,
   numFmtId_  VARCHAR2    := null,
   fontId_    VARCHAR2    := null,
   fillId_    VARCHAR2    := null,
   borderId_  VARCHAR2    := null,
   alignment_ VARCHAR2    := null,
   sheet_     PLS_INTEGER := null )
IS BEGIN
   Cell (
      col_, row_, value_,
      CASE WHEN numFmtId_  IS NOT null THEN numFmt_(numFmtId_) END,
      CASE WHEN fontId_    IS NOT null THEN fonts_(fontId_) END,
      CASE WHEN fillId_    IS NOT null THEN fills_(fillId_) END,
      CASE WHEN borderId_  IS NOT null THEN bdrs_(borderId_) END,
      CASE WHEN alignment_ IS NOT null THEN align_(alignment_) END,
      sheet_
   );
END CellP;

PROCEDURE Cell (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     DATE,
   numFmtId_  PLS_INTEGER  := null,
   fontId_    PLS_INTEGER  := null,
   fillId_    PLS_INTEGER  := null,
   borderId_  PLS_INTEGER  := null,
   alignment_ tp_alignment := null,
   sheet_     PLS_INTEGER  := null )
IS
   num_fmt_id_ PLS_INTEGER := numFmtId_;
   sh_         PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).rows(row_)(col_).value := (value_ - date '1900-03-01' ) + 61;
   IF num_fmt_id_ IS null
      AND not (    workbook.sheets(sh_).col_fmts.exists(col_)
               AND workbook.sheets(sh_).col_fmts(col_).numFmtId IS not null )
      AND not (    workbook.sheets(sh_).row_fmts.exists(row_)
               AND workbook.sheets(sh_).row_fmts(row_).numFmtId IS not null )
   THEN
      num_fmt_id_ := get_numFmt('dd/mm/yyyy');
   END IF;
   workbook.sheets(sh_).rows(row_)(col_).style := get_XfId (
      sh_, col_, row_, num_fmt_id_, fontId_, fillId_, borderId_, alignment_
   );
END Cell;

PROCEDURE CellP (
   col_       PLS_INTEGER,
   row_       PLS_INTEGER,
   value_     DATE,
   numFmtId_  VARCHAR2    := null,
   fontId_    VARCHAR2    := null,
   fillId_    VARCHAR2    := null,
   borderId_  VARCHAR2    := null,
   alignment_ VARCHAR2    := null,
   sheet_     PLS_INTEGER := null )
IS BEGIN
   Cell (
      col_, row_, value_,
      CASE WHEN numFmtId_  IS NOT null THEN numFmt_(numFmtId_) END,
      CASE WHEN fontId_    IS NOT null THEN fonts_(fontId_) END,
      CASE WHEN fillId_    IS NOT null THEN fills_(fillId_) END,
      CASE WHEN borderId_  IS NOT null THEN bdrs_(borderId_) END,
      CASE WHEN alignment_ IS NOT null THEN align_(alignment_) END,
      sheet_
   );
END CellP;

PROCEDURE Query_Date_Cell (
   col_   PLS_INTEGER,
   row_   PLS_INTEGER,
   value_ DATE,
   sheet_ PLS_INTEGER := null,
   XfId_  VARCHAR2 )
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   Cell (col_, row_, value_, 0, sheet_ => sheet_);
   workbook.sheets(sh_).rows(row_)(col_).style := XfId_;
END Query_Date_Cell;

PROCEDURE Condition_Color_Col (
   col_   PLS_INTEGER,
   sheet_ PLS_INTEGER := null )
IS
   sh_        PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
   first_row_ PLS_INTEGER := workbook.sheets(sh_).rows.FIRST;
   last_row_  PLS_INTEGER := workbook.sheets(sh_).rows.LAST;
   str_ix_    PLS_INTEGER;
   str_val_   VARCHAR2(50);
   XfId_      PLS_INTEGER;
   num_fmt_   PLS_INTEGER;
   font_id_   PLS_INTEGER;
   border_id_ PLS_INTEGER;
   align_     tp_alignment;

BEGIN

   FOR r_ IN first_row_ .. last_row_ LOOP

      str_ix_  := workbook.sheets(sh_).rows(r_)(col_).value;
      str_val_ := substr (workbook.str_ind(str_ix_), 1, 50);

      IF fills_.exists(str_val_) THEN

         XfId_ := Extract_Id_From_Style (workbook.sheets(sh_).rows(r_)(col_).style);

         IF XfId_ IS null THEN
            workbook.sheets(sh_).rows(r_)(col_).style := 't="s" ' || get_XfId (
               sh_, col_, r_, fillId_ => fills_(str_val_)
            );
         ELSE
            num_fmt_          := workbook.cellXfs(XfId_).numFmtId;
            font_id_          := workbook.cellXfs(XfId_).fontId;
            border_id_        := workbook.cellXfs(XfId_).borderId;
            align_.vertical   := workbook.cellXfs(XfId_).alignment.vertical;
            align_.horizontal := workbook.cellXfs(XfId_).alignment.horizontal;
            align_.wrapText   := workbook.cellXfs(XfId_).alignment.wrapText;
            workbook.sheets(sh_).rows(r_)(col_).style := 't="s" ' || get_XfId (
               sh_, col_, r_, num_fmt_, font_id_, fills_(str_val_), border_id_, align_
            );
         END IF;

      END IF;

   END LOOP;

END Condition_Color_Col;

PROCEDURE Hyperlink (
   col_   PLS_INTEGER,
   row_   PLS_INTEGER,
   url_   VARCHAR2,
   value_ VARCHAR2    := null,
   sheet_ PLS_INTEGER := null )
IS
   ix_ PLS_INTEGER;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sheet_).rows(row_)(col_).value := add_string(nvl(value_, url_));
   workbook.sheets(sheet_).rows(row_)(col_).style := 't="s" ' || get_XfId(sh_, col_, row_, '', Get_Font('Calibri', theme_ => 10, underline_ => true));
   ix_ := workbook.sheets(sheet_).hyperlinks.count() + 1;
   workbook.sheets(sheet_).hyperlinks(ix_).cell := alfan_col(col_) || row_;
   workbook.sheets(sheet_).hyperlinks(ix_).url := url_;
END Hyperlink;

PROCEDURE Comment (
   col_    PLS_INTEGER,
   row_    PLS_INTEGER,
   text_   VARCHAR2,
   author_ VARCHAR2 := null,
   width_  PLS_INTEGER := 150,
   height_ PLS_INTEGER := 100,
   sheet_  PLS_INTEGER := null )
IS
   ix_ PLS_INTEGER;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   ix_ := workbook.sheets(sh_).comments.count() + 1;
   workbook.sheets(sh_).comments(ix_).row    := row_;
   workbook.sheets(sh_).comments(ix_).column := col_;
   workbook.sheets(sh_).comments(ix_).text   := dbms_xmlgen.convert(text_);
   workbook.sheets(sh_).comments(ix_).author := dbms_xmlgen.convert(author_);
   workbook.sheets(sh_).comments(ix_).width  := width_;
   workbook.sheets(sh_).comments(ix_).height := height_;
END Comment;

PROCEDURE Num_Formula (
   col_           PLS_INTEGER,
   row_           PLS_INTEGER,
   formula_       VARCHAR2,
   default_value_ NUMBER       := null,
   numFmtId_      PLS_INTEGER  := null,
   fontId_        PLS_INTEGER  := null,
   fillId_        PLS_INTEGER  := null,
   borderId_      PLS_INTEGER  := null,
   alignment_     tp_alignment := null,
   sheet_         PLS_INTEGER  := null )
IS
   ix_ PLS_INTEGER := workbook.formulas.count;
   sh_ PLS_INTEGER := nvl (sheet_, workbook.sheets.count());
BEGIN
   workbook.formulas(ix_) := formula_;
   Cell (col_, row_, default_value_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sh_);
   workbook.sheets(sh_).rows(row_)(col_).formula_idx := ix_;
END Num_Formula;

PROCEDURE Str_Formula (
   col_           PLS_INTEGER,
   row_           PLS_INTEGER,
   formula_       VARCHAR2,
   default_value_ VARCHAR2     := null,
   numFmtId_      PLS_INTEGER  := null,
   fontId_        PLS_INTEGER  := null,
   fillId_        PLS_INTEGER  := null,
   borderId_      PLS_INTEGER  := null,
   alignment_     tp_alignment := null,
   sheet_         PLS_INTEGER  := null )
IS
   ix_ PLS_INTEGER := workbook.formulas.count;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.formulas(ix_) := formula_;
   Cell (col_, row_, default_value_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sh_);
   workbook.sheets(sh_).rows(row_)(col_).formula_idx := ix_;
END Str_Formula;

PROCEDURE Mergecells (
   tl_col_ IN PLS_INTEGER, -- top left
   tl_row_ IN PLS_INTEGER,
   br_col_ IN PLS_INTEGER, -- bottom right
   br_row_ IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null )
IS
   ix_   PLS_INTEGER;
   sh_ PLS_INTEGER := nvl (sheet_, workbook.sheets.count());
BEGIN
   ix_ := workbook.sheets(sh_).mergecells.count() + 1;
   workbook.sheets(sh_).mergecells(ix_) :=
      alfan_col(tl_col_) || tl_row_ || ':' || alfan_col(br_col_) || br_row_;
END Mergecells;

PROCEDURE Add_Validation (
   p_type        IN VARCHAR2,
   p_sqref       IN VARCHAR2,
   p_style       IN VARCHAR2    := 'stop', -- stop, warning, information
   p_formula1    IN VARCHAR2    := null,
   p_formula2    IN VARCHAR2    := null,
   p_title       IN VARCHAR2    := null,
   p_prompt      IN VARCHAR     := null,
   p_show_error  IN BOOLEAN     := false,
   p_error_title IN VARCHAR2    := null,
   p_error_txt   IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null )
IS
   ix_     PLS_INTEGER;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   ix_ := workbook.sheets(sh_).validations.count() + 1;
   workbook.sheets(sh_).validations(ix_).type        := p_type;
   workbook.sheets(sh_).validations(ix_).errorstyle  := p_style;
   workbook.sheets(sh_).validations(ix_).sqref       := p_sqref;
   workbook.sheets(sh_).validations(ix_).formula1    := p_formula1;
   workbook.sheets(sh_).validations(ix_).formula2    := p_formula2;
   workbook.sheets(sh_).validations(ix_).error_title := p_error_title;
   workbook.sheets(sh_).validations(ix_).error_txt   := p_error_txt;
   workbook.sheets(sh_).validations(ix_).title       := p_title;
   workbook.sheets(sh_).validations(ix_).prompt      := p_prompt;
   workbook.sheets(sh_).validations(ix_).showerrormessage := p_show_error;
END Add_Validation;

PROCEDURE List_Validation (
   p_sqref_col    IN PLS_INTEGER,
   p_sqref_row    IN PLS_INTEGER,
   p_tl_col       IN PLS_INTEGER, -- top left
   p_tl_row       IN PLS_INTEGER,
   p_br_col       IN PLS_INTEGER, -- bottom right
   p_br_row       IN PLS_INTEGER,
   p_style        IN VARCHAR2    := 'stop', -- stop, warning, information
   p_title        IN VARCHAR2    := null,
   p_prompt       IN VARCHAR     := null,
   p_show_error   IN BOOLEAN     := false,
   p_error_title  IN VARCHAR2    := null,
   p_error_txt    IN VARCHAR2    := null,
   sheet_        IN PLS_INTEGER := null )
IS BEGIN
   Add_Validation (
      p_type        => 'list',
      p_sqref       => alfan_col(p_sqref_col) || p_sqref_row,
      p_style       => lower(p_style),
      p_formula1    => '$' || alfan_col(p_tl_col) || '$' || p_tl_row || ':$' || alfan_col(p_br_col) || '$' || p_br_row,
      p_title       => p_title,
      p_prompt      => p_prompt,
      p_show_error  => p_show_error,
      p_error_title => p_error_title,
      p_error_txt   => p_error_txt,
      sheet_       => sheet_
   );
END List_Validation;

PROCEDURE List_Validation (
   p_sqref_col    IN PLS_INTEGER,
   p_sqref_row    IN PLS_INTEGER,
   p_defined_name IN VARCHAR2,
   p_style        IN VARCHAR2    := 'stop', -- stop, warning, information
   p_title        IN VARCHAR2    := null,
   p_prompt       IN VARCHAR     := null,
   p_show_error   IN BOOLEAN     := false,
   p_error_title  IN VARCHAR2    := null,
   p_error_txt    IN VARCHAR2    := null,
   sheet_        IN PLS_INTEGER := null )
IS BEGIN
   Add_Validation (
      p_type        => 'list',
      p_sqref       => alfan_col(p_sqref_col) || p_sqref_row,
      p_style       => lower(p_style),
      p_formula1    => p_defined_name,
      p_title       => p_title,
      p_prompt      => p_prompt,
      p_show_error  => p_show_error,
      p_error_title => p_error_title,
      p_error_txt   => p_error_txt,
      sheet_       => sheet_
   );
END List_Validation;
--
PROCEDURE Defined_Name (
   tl_col_     PLS_INTEGER, -- top left
   tl_row_     PLS_INTEGER,
   br_col_     PLS_INTEGER, -- bottom right
   br_row_     PLS_INTEGER,
   name_       VARCHAR2,
   sheet_      PLS_INTEGER := null,
   localsheet_ PLS_INTEGER := null )
IS
   ix_   PLS_INTEGER;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   ix_ := workbook.defined_names.count() + 1;
   workbook.defined_names(ix_).name := name_;
   workbook.defined_names(ix_).ref := 'Sheet' || sh_ || '!$' || alfan_col(tl_col_) || '$' || tl_row_ || ':$' || alfan_col(br_col_) || '$' || br_row_;
   workbook.defined_names(ix_).sheet := localsheet_;
END Defined_Name;

PROCEDURE Set_Column_Width (
   col_   PLS_INTEGER,
   width_ NUMBER,
   sheet_ PLS_INTEGER := null )
IS
   w_  NUMBER      := trunc(round(width_*7)*256/7)/256;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).widths(col_) := w_;
END Set_Column_Width;

PROCEDURE Set_Column (
   col_       PLS_INTEGER,
   numFmtId_  PLS_INTEGER  := null,
   fontId_    PLS_INTEGER  := null,
   fillId_    PLS_INTEGER  := null,
   borderId_  PLS_INTEGER  := null,
   alignment_ tp_alignment := null,
   sheet_     PLS_INTEGER  := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).col_fmts(col_).numFmtId  := numFmtId_;
   workbook.sheets(sh_).col_fmts(col_).fontId    := fontId_;
   workbook.sheets(sh_).col_fmts(col_).fillId    := fillId_;
   workbook.sheets(sh_).col_fmts(col_).borderId  := borderId_;
   workbook.sheets(sh_).col_fmts(col_).alignment := alignment_;
END Set_Column;

PROCEDURE Set_Row (
   row_       IN PLS_INTEGER,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null,
   height_    IN NUMBER       := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
   c_  tp_cells;
BEGIN
   workbook.sheets(sh_).row_fmts(row_).numFmtId  := numFmtId_;
   workbook.sheets(sh_).row_fmts(row_).fontId    := fontId_;
   workbook.sheets(sh_).row_fmts(row_).fillId    := fillId_;
   workbook.sheets(sh_).row_fmts(row_).borderId  := borderId_;
   workbook.sheets(sh_).row_fmts(row_).alignment := alignment_;
   workbook.sheets(sh_).row_fmts(row_).height    := trunc(height_*4/3)*3/4;
   IF not workbook.sheets(sh_).rows.exists(row_) THEN
      workbook.sheets(sh_).rows(row_) := c_;
   END IF;
END Set_Row;

PROCEDURE Freeze_Rows (
   nr_rows_ IN PLS_INTEGER := 1,
   sheet_   IN PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).freeze_cols := null;
   workbook.sheets(sh_).freeze_rows := nr_rows_;
END Freeze_Rows;

PROCEDURE Freeze_Cols (
   nr_cols_ IN PLS_INTEGER := 1,
   sheet_   IN PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).freeze_rows := null;
   workbook.sheets(sh_).freeze_cols := nr_cols_;
END Freeze_Cols;

PROCEDURE Freeze_Pane (
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER,
   sheet_ IN PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl (sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).freeze_rows := row_;
   workbook.sheets(sh_).freeze_cols := col_;
END Freeze_Pane;

PROCEDURE Set_Autofilter (
   col_start_ IN PLS_INTEGER := null,
   col_end_   IN PLS_INTEGER := null,
   row_start_ IN PLS_INTEGER := null,
   row_end_   IN PLS_INTEGER := null,
   sheet_     IN PLS_INTEGER := null )
IS
   ix_ PLS_INTEGER := 1;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).autofilters(ix_).column_start := col_start_;
   workbook.sheets(sh_).autofilters(ix_).column_end   := col_end_;
   workbook.sheets(sh_).autofilters(ix_).row_start    := row_start_;
   workbook.sheets(sh_).autofilters(ix_).row_end      := row_end_;
   Defined_Name (col_start_, row_start_, col_end_, row_end_, '_xlnm._FilterDatabase', sh_, sh_-1);
END Set_Autofilter;

PROCEDURE Add1Xml (
   excel_    IN OUT NOCOPY BLOB,
   filename_ VARCHAR2,
   xml_      CLOB )
IS
   tmp_          BLOB;
   dest_offset_  INTEGER := 1;
   src_offset_   INTEGER := 1;
   lang_context_ INTEGER := Dbms_Lob.DEFAULT_LANG_CTX;
   warning_      INTEGER;
BEGIN
   Dbms_Lob.CreateTemporary (tmp_, true);
   Dbms_Lob.ConvertToBlob (
      tmp_, xml_, Dbms_Lob.LobMaxSize, dest_offset_, src_offset_,
      nls_charset_id('AL32UTF8'), lang_context_, warning_
   );
   Add1File (excel_, filename_, tmp_);
   Dbms_Lob.freetemporary(tmp_);
END Add1Xml;
--
FUNCTION Finish RETURN BLOB
IS
   excel_   BLOB;
   yyy_     BLOB;
   xxx_     CLOB;
   tmp_     VARCHAR2(32767 char);
   c_       NUMBER;
   h_       NUMBER;
   w_       NUMBER;
   cw_      NUMBER;
   s        PLS_INTEGER;
   row_ix_  PLS_INTEGER;
   col_ix_  PLS_INTEGER;
   col_min_ PLS_INTEGER;
   col_max_ PLS_INTEGER;

BEGIN
  dbms_lob.createtemporary(excel_, true);
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
  s := workbook.sheets.first;
  WHILE s IS not null LOOP
    xxx_ := xxx_ || ( '
<Override PartName="/xl/worksheets/sheet' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' );
    s := workbook.sheets.next(s);
  END LOOP;
  xxx_ := xxx_ || '
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
  s := workbook.sheets.first;
  WHILE s IS not null LOOP
    IF workbook.sheets( s ).comments.count() > 0 THEN
      xxx_ := xxx_ || ( '
<Override PartName="/xl/comments' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' );
    END IF;
    s := workbook.sheets.next( s );
  END LOOP;
  xxx_ := xxx_ || '
</Types>';
  add1xml (excel_, '[Content_Types].xml', xxx_);
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>' || sys_context( 'userenv', 'os_user' ) || '</dc:creator>
<dc:description>Build by version:' || VERSION_ || '</dc:description>
<cp:lastModifiedBy>' || sys_context( 'userenv', 'os_user' ) || '</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:modified>
</cp:coreProperties>';
  add1xml (excel_, 'docProps/core.xml', xxx_);
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
<vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>' || workbook.sheets.count() || '</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="' || workbook.sheets.count() || '" baseType="lpstr">';
  s := workbook.sheets.first;
  WHILE s IS not null LOOP
    xxx_ := xxx_ || ( '
<vt:lpstr>' || workbook.sheets(s).name || '</vt:lpstr>' );
    s := workbook.sheets.next(s);
  END LOOP;
  xxx_ := xxx_ || '</vt:vector>
</TitlesOfParts>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>14.0300</AppVersion>
</Properties>';
  add1xml( excel_, 'docProps/app.xml', xxx_ );
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
  add1xml( excel_, '_rels/.rels', xxx_ );
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
  IF workbook.numFmts.count() > 0 THEN
    xxx_ := xxx_ || ( '<numFmts count="' || workbook.numFmts.count() || '">' );
    FOR n IN 1 .. workbook.numFmts.count() LOOP
      xxx_ := xxx_ || ( '<numFmt numFmtId="' || workbook.numFmts(n).numFmtId || '" formatCode="' || workbook.numFmts(n).formatCode || '"/>' );
    END LOOP;
    xxx_ := xxx_ || '</numFmts>';
  END IF;
  xxx_ := xxx_ || ( '<fonts count="' || workbook.fonts.count() || '" x14ac:knownFonts="1">' );
  FOR f IN 0 .. workbook.fonts.count() - 1 LOOP
    xxx_ := xxx_ || ( '<font>' ||
      CASE WHEN workbook.fonts(f).bold THEN '<b/>' END ||
      CASE WHEN workbook.fonts(f).italic THEN '<i/>' END ||
      CASE WHEN workbook.fonts(f).underline THEN '<u/>' END ||
'<sz val="' || to_char( workbook.fonts(f).fontsize, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )  || '"/>
<color ' || CASE WHEN workbook.fonts(f).rgb IS not null
              THEN 'rgb="' || workbook.fonts( f ).rgb
              ELSE 'theme="' || workbook.fonts(f).theme
            END || '"/>
<name val="' || workbook.fonts(f).name || '"/>
<family val="' || workbook.fonts(f).family || '"/>
<scheme val="none"/>
</font>' );
  END LOOP;
  xxx_ := xxx_ || ( '</fonts>
<fills count="' || workbook.fills.count() || '">' );
  FOR f IN 0 .. workbook.fills.count() - 1 LOOP
    xxx_ := xxx_ || ( '<fill><patternFill patternType="' || workbook.fills(f).patternType || '">' ||
      CASE WHEN workbook.fills(f).fgRGB IS not null THEN '<fgColor rgb="' || workbook.fills(f).fgRGB || '"/>' END ||
      CASE WHEN workbook.fills(f).bgRGB IS not null THEN '<bgColor rgb="' || workbook.fills(f).bgRGB || '"/>' END ||
          '</patternFill></fill>' );
  END LOOP;
  xxx_ := xxx_ || ( '</fills>
<borders count="' || workbook.borders.count() || '">' );
  FOR b IN 0 .. workbook.borders.count() - 1 LOOP
    xxx_ := xxx_ || ('<border>' ||
      CASE WHEN workbook.borders(b).left   IS null THEN '<left/>'   ELSE '<left style="'   || workbook.borders(b).left   || '"/>' END ||
      CASE WHEN workbook.borders(b).right  IS null THEN '<right/>'  ELSE '<right style="'  || workbook.borders(b).right  || '"/>' END ||
      CASE WHEN workbook.borders(b).top    IS null THEN '<top/>'    ELSE '<top style="'    || workbook.borders(b).top    || '"/>' END ||
      CASE WHEN workbook.borders(b).bottom IS null THEN '<bottom/>' ELSE '<bottom style="' || workbook.borders(b).bottom || '"/>' END ||
      '</border>'
    );
  END LOOP;
  xxx_ := xxx_ || ( '</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="' || ( workbook.cellXfs.count() + 1 ) || '">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' );
  FOR x IN 1 .. workbook.cellXfs.count() LOOP
    xxx_ := xxx_ || ( '<xf numFmtId="' || workbook.cellXfs(x).numFmtId || '" fontId="' || workbook.cellXfs(x).fontId || '" fillId="' || workbook.cellXfs(x).fillId || '" borderId="' || workbook.cellXfs(x).borderId || '">' );
    IF (    workbook.cellXfs(x).alignment.horizontal IS not null
         OR workbook.cellXfs(x).alignment.vertical IS not null
         OR workbook.cellXfs(x).alignment.wrapText )
    THEN
      xxx_ := xxx_ || ( '<alignment' ||
        CASE WHEN workbook.cellXfs(x).alignment.horizontal IS not null THEN ' horizontal="' || workbook.cellXfs(x).alignment.horizontal || '"' END ||
        CASE WHEN workbook.cellXfs(x).alignment.vertical IS not null THEN ' vertical="' || workbook.cellXfs(x).alignment.vertical || '"' END ||
        CASE WHEN workbook.cellXfs(x).alignment.wrapText THEN ' wrapText="true"' END || '/>' );
    END IF;
    xxx_ := xxx_ || '</xf>';
  END LOOP;
  xxx_ := xxx_ || ( '</cellXfs>
<cellStyles count="1">
  <cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
<extLst>
<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
</ext>
</extLst>
</styleSheet>' );
  add1xml( excel_, 'xl/styles.xml', xxx_ );
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr date1904="false" defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
  s := workbook.sheets.first;
  WHILE s IS not null LOOP
    xxx_ := xxx_ || ('
<sheet name="' || workbook.sheets(s).name || '" sheetId="' || s || '" r:id="rId' || ( 9 + s ) || '"/>' );
    s := workbook.sheets.next(s);
  END LOOP;
  xxx_ := xxx_ || '</sheets>';
  IF workbook.defined_names.count() > 0 THEN
    xxx_ := xxx_ || '<definedNames>';
    FOR s IN 1 .. workbook.defined_names.count() LOOP
      xxx_ := xxx_ || ('
<definedName name="' || workbook.defined_names(s).name || '"' ||
        CASE WHEN workbook.defined_names( s ).sheet IS not null THEN ' localSheetId="' || to_char(workbook.defined_names(s).sheet) || '"' END ||
        '>' || workbook.defined_names( s ).ref || '</definedName>');
    END LOOP;
    xxx_ := xxx_ || '</definedNames>';
  END IF;
  xxx_ := xxx_ || '<calcPr calcId="144525"/></workbook>';
  add1xml( excel_, 'xl/workbook.xml', xxx_ );
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Cambria"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0">
      <a:schemeClr val="phClr">
        <a:shade val="51000"/>
        <a:satMod val="130000"/>
      </a:schemeClr>
    </a:gs>
    <a:gs pos="80000">
      <a:schemeClr val="phClr">
        <a:shade val="93000"/>
        <a:satMod val="130000"/>
      </a:schemeClr>
    </a:gs>
    <a:gs pos="100000">
      <a:schemeClr val="phClr">
        <a:shade val="94000"/>
        <a:satMod val="135000"/>
      </a:schemeClr>
    </a:gs>
  </a:gsLst>
  <a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>';
  add1xml (excel_, 'xl/theme/theme1.xml', xxx_);
  s := workbook.sheets.first;
  WHILE s IS not null LOOP
    col_min_ := 16384;
    col_max_ := 1;
    row_ix_ := workbook.sheets(s).rows.first();
    WHILE row_ix_ IS not null LOOP
      col_min_ := least(col_min_, workbook.sheets(s).rows(row_ix_).first());
      col_max_ := greatest(col_max_, workbook.sheets(s).rows(row_ix_).last());
      row_ix_ := workbook.sheets(s).rows.next(row_ix_);
    END LOOP;
    addtxt2utf8blob_init(yyy_);
    addtxt2utf8blob ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' ||
CASE WHEN workbook.sheets(s).tabcolor IS not null THEN '<sheetPr><tabColor rgb="' || workbook.sheets(s).tabcolor || '"/></sheetPr>' end ||
'<dimension ref="' || alfan_col(col_min_) || workbook.sheets(s).rows.first() || ':' || alfan_col(col_max_) || workbook.sheets(s).rows.last() || '"/>
<sheetViews>
<sheetView' || CASE WHEN s = 1 THEN ' tabSelected="1"' END || ' workbookViewId="0">', yyy_);
    IF workbook.sheets(s).freeze_rows > 0 AND workbook.sheets(s).freeze_cols > 0 THEN
      addtxt2utf8blob (
        '<pane xSplit="' || workbook.sheets(s).freeze_cols || '" '
        || 'ySplit="' || workbook.sheets(s).freeze_rows || '" '
        || 'topLeftCell="' || alfan_col(workbook.sheets(s).freeze_cols+1) || (workbook.sheets(s).freeze_rows+1) || '" '
        || 'activePane="bottomLeft" state="frozen"/>',
        yyy_
      );
    ELSE
      IF workbook.sheets(s).freeze_rows > 0 THEN
        addtxt2utf8blob (
          '<pane ySplit="' || workbook.sheets(s).freeze_rows || '" topLeftCell="A' ||
            (workbook.sheets(s).freeze_rows+1) || '" activePane="bottomLeft" state="frozen"/>',
          yyy_
        );
      END IF;
      IF workbook.sheets(s).freeze_cols > 0 THEN
        addtxt2utf8blob (
          '<pane xSplit="' || workbook.sheets(s).freeze_cols || '" topLeftCell="' ||
          alfan_col( workbook.sheets( s ).freeze_cols + 1 ) ||
          '1" activePane="bottomLeft" state="frozen"/>',
          yyy_
        );
      END IF;
    END IF;
    addtxt2utf8blob ('</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>', yyy_);
    IF workbook.sheets(s).widths.count() > 0 THEN
      addtxt2utf8blob ('<cols>', yyy_);
      col_ix_ := workbook.sheets(s).widths.first();
      WHILE col_ix_ IS not null LOOP
        addtxt2utf8blob ('<col min="' || col_ix_ || '" max="' || col_ix_ || '" width="' || to_char(workbook.sheets(s).widths(col_ix_), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" customWidth="1"/>', yyy_);
        col_ix_ := workbook.sheets(s).widths.next(col_ix_);
      END LOOP;
      addtxt2utf8blob('</cols>', yyy_);
    END IF;
    addtxt2utf8blob('<sheetData>', yyy_);
    row_ix_ := workbook.sheets(s).rows.first();
    WHILE row_ix_ IS not null LOOP
      IF workbook.sheets(s).row_fmts.exists(row_ix_) AND workbook.sheets(s).row_fmts(row_ix_).height IS not null THEN
          addtxt2utf8blob( '<row r="' || row_ix_ || '" spans="' || col_min_ || ':' || col_max_ || '" customHeight="1" ht="'
                         || to_char( workbook.sheets(s).row_fmts(row_ix_).height, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" >', yyy_ );
      ELSE
        addtxt2utf8blob( '<row r="' || row_ix_ || '" spans="' || col_min_ || ':' || col_max_ || '">', yyy_ );
      END IF;
      col_ix_ := workbook.sheets(s).rows(row_ix_).first();
      WHILE col_ix_ IS not null LOOP
        IF workbook.sheets(s).rows(row_ix_)(col_ix_).formula_idx IS null THEN
          tmp_ := null;
        ELSE
          tmp_ := '<f>' || workbook.formulas(workbook.sheets(s).rows(row_ix_)(col_ix_).formula_idx) || '</f>';
        END IF;
        addtxt2utf8blob ('<c r="' || alfan_col(col_ix_) || row_ix_ || '"'
          || ' ' || workbook.sheets(s).rows(row_ix_)(col_ix_).style
          || '>' || tmp_ || '<v>'
          || to_char(workbook.sheets(s).rows(row_ix_)(col_ix_).value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )
          || '</v></c>', yyy_
        );
        col_ix_ := workbook.sheets(s).rows(row_ix_).next(col_ix_);
      END LOOP;
      addtxt2utf8blob( '</row>', yyy_ );
      row_ix_ := workbook.sheets( s ).rows.next(row_ix_);
    END LOOP;
    addtxt2utf8blob( '</sheetData>', yyy_ );
    FOR a IN 1 ..  workbook.sheets(s).autofilters.count() LOOP
      addtxt2utf8blob( '<autoFilter ref="' ||
        alfan_col( nvl( workbook.sheets(s).autofilters(a).column_start, col_min_ ) ) ||
        nvl( workbook.sheets(s).autofilters(a).row_start, workbook.sheets(s).rows.first() ) || ':' ||
        alfan_col(coalesce( workbook.sheets(s).autofilters(a).column_end, workbook.sheets( s ).autofilters(a).column_start, col_max_)) ||
        nvl(workbook.sheets(s).autofilters(a).row_end, workbook.sheets(s).rows.last()) || '"/>', yyy_);
    END LOOP;
    IF workbook.sheets(s).mergecells.count() > 0 THEN
      addtxt2utf8blob( '<mergeCells count="' || to_char(workbook.sheets(s).mergecells.count()) || '">', yyy_);
      FOR m IN 1 ..  workbook.sheets(s).mergecells.count() LOOP
        addtxt2utf8blob( '<mergeCell ref="' || workbook.sheets( s ).mergecells( m ) || '"/>', yyy_);
      END LOOP;
      addtxt2utf8blob('</mergeCells>', yyy_);
    END IF;
--
    IF workbook.sheets(s).validations.count() > 0 THEN
      addtxt2utf8blob( '<dataValidations count="' || to_char( workbook.sheets( s ).validations.count() ) || '">', yyy_ );
      FOR m IN 1 ..  workbook.sheets( s ).validations.count() LOOP
        addtxt2utf8blob ('<dataValidation' ||
            ' type="' || workbook.sheets(s).validations(m).type || '"' ||
            ' errorStyle="' || workbook.sheets(s).validations(m).errorstyle || '"' ||
            ' allowBlank="' || CASE WHEN nvl(workbook.sheets(s).validations(m).allowBlank, true) THEN '1' ELSE '0' END || '"' ||
            ' sqref="' || workbook.sheets(s).validations(m).sqref || '"', yyy_ );
        IF workbook.sheets( s ).validations(m).prompt IS not null THEN
          addtxt2utf8blob(' showInputMessage="1" prompt="' || workbook.sheets(s).validations(m).prompt || '"', yyy_);
          IF workbook.sheets(s).validations(m).title IS not null THEN
            addtxt2utf8blob( ' promptTitle="' || workbook.sheets(s).validations(m).title || '"', yyy_);
          END IF;
        END IF;
        IF workbook.sheets(s).validations(m).showerrormessage THEN
          addtxt2utf8blob( ' showErrorMessage="1"', yyy_);
          IF workbook.sheets(s).validations(m).error_title IS not null THEN
            addtxt2utf8blob( ' errorTitle="' || workbook.sheets(s).validations(m).error_title || '"', yyy_);
          END IF;
          IF workbook.sheets(s).validations(m).error_txt IS not null THEN
            addtxt2utf8blob(' error="' || workbook.sheets(s).validations(m).error_txt || '"', yyy_);
          END IF;
        END IF;
        addtxt2utf8blob( '>', yyy_ );
        IF workbook.sheets(s).validations(m).formula1 IS not null THEN
          addtxt2utf8blob ('<formula1>' || workbook.sheets(s).validations(m).formula1 || '</formula1>', yyy_);
        END IF;
        IF workbook.sheets(s).validations(m).formula2 IS not null THEN
          addtxt2utf8blob ('<formula2>' || workbook.sheets(s).validations(m).formula2 || '</formula2>', yyy_);
        END IF;
        addtxt2utf8blob ('</dataValidation>', yyy_);
      END LOOP;
      addtxt2utf8blob ('</dataValidations>', yyy_);
    END IF;

    IF workbook.sheets(s).hyperlinks.count() > 0 THEN
      addtxt2utf8blob ('<hyperlinks>', yyy_);
      FOR h IN 1 ..  workbook.sheets( s ).hyperlinks.count() LOOP
        addtxt2utf8blob ('<hyperlink ref="' || workbook.sheets(s).hyperlinks(h).cell || '" r:id="rId' || h || '"/>', yyy_);
      END LOOP;
      addtxt2utf8blob ('</hyperlinks>', yyy_);
    END IF;
    addtxt2utf8blob( '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', yyy_ );
    IF workbook.sheets(s).comments.count() > 0 THEN
      addtxt2utf8blob( '<legacyDrawing r:id="rId' || ( workbook.sheets(s).hyperlinks.count() + 1 ) || '"/>', yyy_ );
    END IF;

    addtxt2utf8blob( '</worksheet>', yyy_ );
    addtxt2utf8blob_finish( yyy_ );
    add1file( excel_, 'xl/worksheets/sheet' || s || '.xml', yyy_ );
    IF workbook.sheets(s).hyperlinks.count() > 0 OR workbook.sheets(s).comments.count() > 0 THEN
      xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
      IF workbook.sheets(s).comments.count() > 0 THEN
        xxx_ := xxx_ || ( '<Relationship Id="rId' || ( workbook.sheets(s).hyperlinks.count() + 2 ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' || s || '.xml"/>' );
        xxx_ := xxx_ || ( '<Relationship Id="rId' || ( workbook.sheets(s).hyperlinks.count() + 1 ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' || s || '.vml"/>' );
      END IF;
      FOR h IN 1 ..  workbook.sheets( s ).hyperlinks.count() LOOP
        xxx_ := xxx_ || ( '<Relationship Id="rId' || h || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' || workbook.sheets( s ).hyperlinks( h ).url || '" TargetMode="External"/>' );
      END LOOP;
      xxx_ := xxx_ || '</Relationships>';
      add1xml( excel_, 'xl/worksheets/_rels/sheet' || s || '.xml.rels', xxx_ );
    END IF;

    IF workbook.sheets(s).comments.count() > 0 THEN
      DECLARE
        cnt PLS_INTEGER;
        author_ind tp_author;
        -- col_ix_ := workbook.sheets(s).widths.next(col_ix_);
      BEGIN
        authors.delete();
        FOR c IN 1 .. workbook.sheets(s).comments.count() LOOP
          authors(workbook.sheets(s).comments(c).author) := 0;
        END LOOP;
        xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>';
        cnt := 0;
        author_ind := authors.first();
        WHILE author_ind IS not null OR authors.next(author_ind) IS not null LOOP
          authors( author_ind ) := cnt;
          xxx_ := xxx_ || ( '<author>' || author_ind || '</author>' );
          cnt := cnt + 1;
          author_ind := authors.next(author_ind);
        END LOOP;
      END;
      xxx_ := xxx_ || '</authors><commentList>';
      FOR c IN 1 .. workbook.sheets( s ).comments.count() LOOP
        xxx_ := xxx_ || ( '<comment ref="' || alfan_col( workbook.sheets(s).comments(c).column ) ||
           to_char (workbook.sheets(s).comments(c).row || '" authorId="' || authors(workbook.sheets(s).comments(c).author ) ) || '">
<text>');
        IF workbook.sheets(s).comments(c).author IS not null THEN
          xxx_ := xxx_ || ( '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
             workbook.sheets(s).comments(c).author || ':</t></r>' );
        END IF;
        xxx_ := xxx_ || ( '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
             CASE WHEN workbook.sheets(s).comments(c).author IS not null THEN '
' END || workbook.sheets(s).comments(c).text || '</t></r></text></comment>' );
      END LOOP;
      xxx_ := xxx_ || '</commentList></comments>';
      add1xml (excel_, 'xl/comments' || s || '.xml', xxx_);
      xxx_ := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
      FOR c IN 1 .. workbook.sheets(s).comments.count() LOOP
        xxx_ := xxx_ || ( '<v:shape id="_x0000_s' || to_char(c) || '" type="#_x0000_t202"
style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' || to_char( c ) || ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' );
        w_ := workbook.sheets(s).comments(c).width;
        c_ := 1;
        LOOP
          IF workbook.sheets(s).widths.exists(workbook.sheets(s).comments(c).column+c_) THEN
            cw_ := 256 * workbook.sheets(s).widths(workbook.sheets(s).comments(c).column+c_);
            cw_ := trunc((cw_+18)/256*7); -- assume default 11 point Calibri
          ELSE
            cw_ := 64;
          END IF;
          EXIT WHEN w_ < cw_;
          c_ := c_ + 1;
          w_ := w_ - cw_;
        END LOOP;
        h_ := workbook.sheets(s).comments(c).height;
        xxx_ := xxx_ || ( '<x:Anchor>' || workbook.sheets(s).comments(c).column || ',15,' ||
            workbook.sheets(s).comments(c).row || ',30,' ||
            (workbook.sheets(s).comments(c).column+c_-1) || ',' || round(w_) || ',' ||
            (workbook.sheets(s).comments(c).row + 1 + trunc(h_/20)) || ',' || mod(h_, 20) || '</x:Anchor>' );
        xxx_ := xxx_ || ( '<x:AutoFill>False</x:AutoFill><x:Row>' ||
            (workbook.sheets(s).comments(c).row-1) || '</x:Row><x:Column>' ||
            (workbook.sheets(s).comments(c).column-1) || '</x:Column></x:ClientData></v:shape>' );
      END LOOP;
      xxx_ := xxx_ || '</xml>';
      add1xml( excel_, 'xl/drawings/vmlDrawing' || s || '.vml', xxx_ );
    END IF;
--
    s := workbook.sheets.next(s);
  END LOOP;
  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
  s := workbook.sheets.first;
  WHILE s IS not null LOOP
    xxx_ := xxx_ || ( '
<Relationship Id="rId' || (9+s) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || s || '.xml"/>' );
    s := workbook.sheets.next(s);
  END LOOP;
  xxx_ := xxx_ || '</Relationships>';
  add1xml (excel_, 'xl/_rels/workbook.xml.rels', xxx_);
  addtxt2utf8blob_init(yyy_);
  addtxt2utf8blob (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || workbook.str_cnt || '" uniqueCount="' || workbook.strings.count() || '">',
    yyy_
  );
  FOR i IN 0 .. workbook.str_ind.count() - 1 LOOP
    addtxt2utf8blob (
       '<si><t xml:space="preserve">' ||
       dbms_xmlgen.convert(substr(workbook.str_ind(i), 1, 32000)) || '</t></si>', yyy_
    );
  END LOOP;
  addtxt2utf8blob( '</sst>', yyy_);
  addtxt2utf8blob_finish(yyy_);
  add1file(excel_, 'xl/sharedStrings.xml', yyy_);
  finish_zip(excel_);
  Clear_Workbook;
  RETURN excel_;
END Finish;

PROCEDURE Save (
   directory_ IN VARCHAR2,
   filename_  IN VARCHAR2 )
IS BEGIN
   Blob2File (Finish, directory_, filename_);
END Save;

PROCEDURE Save (
   xl_blob_   IN BLOB,
   directory_ IN VARCHAR2,
   filename_  IN VARCHAR2 )
IS BEGIN
   Blob2File (xl_blob_, directory_, filename_);
END Save;

PROCEDURE Query2Sheet (
   col_count_   IN OUT PLS_INTEGER,
   row_count_   IN OUT PLS_INTEGER,
   cur_         IN OUT INTEGER,
   col_headers_ IN BOOLEAN     := true,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() )
IS

   TYPE tp_XfIds IS TABLE OF VARCHAR2(50) INDEX BY PLS_INTEGER;

   sh_           PLS_INTEGER := sheet_;
   desc_tab_     dbms_sql.desc_tab2;
   d_tab_        dbms_sql.date_table;
   n_tab_        dbms_sql.number_table;
   v_tab_        dbms_sql.varchar2_table;
   data_len_     NUMBER;
   bulk_sz_      PLS_INTEGER := 200;
   rows_fetched_ INTEGER;
   offset_       PLS_INTEGER;
   useXf_bkp_    BOOLEAN := g_useXf_;
   XfIds_        tp_XfIds;
   widths_       tp_widths; --TYPE tp_widths is table of NUMBER index by PLS_INTEGER;
   ix_           NUMBER;

BEGIN

   setUseXf (useXf_); -- originally "true"

   IF sheet_ IS null THEN
      sh_ := New_Sheet;
   END IF;

   Dbms_Sql.Describe_Columns2 (cur_, col_count_, desc_tab_);

   FOR col_ IN 1 .. col_count_ LOOP
      IF col_headers_ THEN
         Cell (
            col_, 1, desc_tab_(col_).col_name, sheet_ => sh_,
            fontId_ => hdr_font_, fillId_ => hdr_fill_
         );
      END IF;
      CASE
         -- Codes for various forms of number (float, number, binary_double)
         WHEN desc_tab_(col_).col_type IN (2, 100, 101) THEN
            dbms_sql.define_array (cur_, col_, n_tab_, bulk_sz_, 1);
         -- Codes for DATE + TIMESTAMP types (with and without time-zone detail)
         WHEN desc_tab_(col_).col_type IN (12, 178, 179, 180, 181, 231) THEN
            dbms_sql.define_array (cur_, col_, d_tab_, bulk_sz_, 1);
            XfIds_(col_) := get_XfId (sh_, col_, null, get_numFmt('dd/mm/yyyy'));
         -- Codes for CHAR + VARCHAR types
         WHEN desc_tab_(col_).col_type IN (1, 8, 9, 96, 112) THEN
            dbms_sql.define_array (cur_, col_, v_tab_, bulk_sz_, 1);
         -- Other stuff (like BLOBs) we can't easily encode into Excel, so we ignore!
         ELSE
            null;
      END CASE;
      widths_(col_) := 8;
   END LOOP;

   offset_    := CASE WHEN col_headers_ THEN 2 ELSE 1 END;
   row_count_ := 0;

   LOOP
      rows_fetched_ := dbms_sql.fetch_rows(cur_);
      row_count_    := row_count_ + rows_fetched_;
      IF rows_fetched_ > 0 THEN
         FOR col_ IN 1 .. col_count_ LOOP
            CASE
               WHEN desc_tab_(col_).col_type IN (2, 100, 101) THEN
                  -- Numbers
                  Dbms_Sql.Column_Value (cur_, col_, n_tab_);
                  FOR i_ IN 0 .. rows_fetched_ - 1 LOOP
                     IF n_tab_(i_+n_tab_.first()) IS NOT null THEN
                        Cell (
                           col_      => col_,
                           row_      => offset_+i_,
                           value_    => n_tab_(i_+n_tab_.first()),
                           numFmtId_ => CASE WHEN col_fmts_.exists(col_) THEN col_fmts_(col_) END,
                           sheet_    => sh_
                        );
                     END IF;
                  END LOOP;
                  n_tab_.delete;
               WHEN desc_tab_(col_).col_type IN (12, 178, 179, 180, 181, 231) THEN
                  -- Dates
                  Dbms_Sql.Column_Value(cur_, col_, d_tab_);
                  FOR i_ IN 0 .. rows_fetched_ - 1 LOOP
                     IF d_tab_(i_+d_tab_.first()) IS NOT null THEN
                        IF g_useXf_ THEN
                           Cell (col_, offset_+i_, d_tab_(i_+d_tab_.first()), sheet_ => sh_);
                        ELSE
                           Query_Date_Cell(col_, offset_+i_, d_tab_(i_+d_tab_.first()), sh_, XfIds_(col_));
                        END IF;
                        widths_(col_) := 12;
                     END IF;
                  END LOOP;
                  d_tab_.delete;
               WHEN desc_tab_(col_).col_type IN (1, 8, 9, 96, 112) THEN
                  -- Text
                  Dbms_Sql.Column_Value (cur_, col_, v_tab_);
                  FOR i_ IN 0 .. rows_fetched_-1 LOOP
                     IF v_tab_(i_+v_tab_.first()) IS NOT null THEN
                        Cell (col_, offset_+i_, v_tab_(i_+v_tab_.first()), sheet_ => sh_);
                        data_len_ := length(v_tab_(i_+v_tab_.first()));
                        widths_(col_) := least (greatest(widths_(col_),data_len_), 60);
                     END IF;
                  END LOOP;
                  v_tab_.delete;
               ELSE
                  null;
            END CASE;
         END LOOP;
      END IF;
      EXIT WHEN rows_fetched_ != bulk_sz_;
      offset_ := offset_ + rows_fetched_;
   END LOOP; -- loop for each column in the result set

   -- set column widths
   ix_ := widths_.first;
   WHILE ix_ IS not null LOOP
      Set_Column_Width (ix_, widths_(ix_), sh_);
      ix_ := widths_.next(ix_);
   END LOOP;

   Dbms_Sql.Close_Cursor (cur_);
   setUseXf (useXf_bkp_);

EXCEPTION
   WHEN others THEN
      IF dbms_sql.is_open (cur_) THEN
         dbms_sql.close_cursor (cur_);
      END IF;
      setUseXf(useXf_);
END Query2Sheet;

PROCEDURE Do_Binding (
   cur_   IN OUT INTEGER,
   binds_ IN OUT NOCOPY bind_arr )
IS
   bind_id_ VARCHAR2(50) := binds_.first;
BEGIN
   LOOP
      EXIT WHEN bind_id_ IS null;
      CASE binds_(bind_id_).datatype
         WHEN 'STRING' THEN Dbms_Sql.Bind_Variable (cur_, bind_id_, binds_(bind_id_).s_val);
         WHEN 'NUMBER' THEN Dbms_Sql.Bind_Variable (cur_, bind_id_, binds_(bind_id_).n_val);
         WHEN 'DATE'   THEN Dbms_Sql.Bind_Variable (cur_, bind_id_, binds_(bind_id_).d_val);
      END CASE;
      bind_id_ := binds_.next(bind_id_);
   END LOOP;
END Do_Binding;

-- Query2Sheet() => Using SQL, with binding
PROCEDURE Query2Sheet (
   col_count_   IN OUT PLS_INTEGER,
   row_count_   IN OUT PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() )
IS
   cur_ INTEGER := Dbms_Sql.Open_Cursor;
   res_ INTEGER;
BEGIN
   Dbms_Sql.Parse (cur_, sql_, dbms_sql.native);
   Do_Binding (cur_, binds_);
   res_ := Dbms_Sql.Execute(cur_);
   SELECT 1 INTO res_ FROM dual; --avoid compiler warning
   Query2Sheet (
      col_count_, row_count_, cur_, col_headers_,
      sheet_, UseXf_, hdr_font_, hdr_fill_, col_fmts_
   );
   IF directory_ IS NOT null AND filename_ IS NOT null THEN
      Save (directory_, filename_);
   END IF;
END Query2Sheet;

-- Query2Sheet() => Using SQL, no binding
PROCEDURE Query2Sheet (
   col_count_   IN OUT PLS_INTEGER,
   row_count_   IN OUT PLS_INTEGER,
   sql_         IN VARCHAR2,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() )
IS
   binds_ bind_arr := bind_arr();
BEGIN
   Query2Sheet (
      col_count_, row_count_, sql_, binds_,
      col_headers_, directory_, filename_, sheet_,
      useXf_, hdr_font_, hdr_fill_, col_fmts_
   );
END Query2Sheet;

-- Query2Sheet() => Using REFCURSOR
PROCEDURE Query2Sheet (
   col_count_   IN OUT PLS_INTEGER,
   row_count_   IN OUT PLS_INTEGER,
   rc_          IN OUT SYS_REFCURSOR,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() )
IS
   cur_ INTEGER := dbms_sql.to_cursor_number (rc_);
BEGIN
   Query2Sheet (
      col_count_, row_count_, cur_, col_headers_,
      sheet_, useXf_, hdr_font_, hdr_fill_, col_fmts_
   );
   IF directory_ IS NOT null AND filename_ IS NOT null THEN
      Save (directory_, filename_);
   END IF;
END Query2Sheet;

PROCEDURE Query2SheetAndAutofilter ( -- with Binds
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() )
IS
   col_count_ NUMBER;
   row_count_ NUMBER;
BEGIN
   Query2Sheet (
      col_count_   => col_count_,
      row_count_   => row_count_,
      sql_         => sql_,
      binds_       => binds_,
      col_headers_ => col_headers_,
      sheet_       => sheet_,
      useXf_       => useXf_,
      hdr_font_    => hdr_font_,
      hdr_fill_    => hdr_fill_,
      col_fmts_    => col_fmts_
   );
   Set_Autofilter (1, col_count_, 1, row_count_, sheet_);
   IF directory_ IS NOT null AND filename_ IS NOT null THEN
      Save (directory_, filename_);
   END IF;
END Query2SheetAndAutofilter;

PROCEDURE Query2SheetAndAutofilter ( -- no Binds
   sql_         IN VARCHAR2,
   col_headers_ IN BOOLEAN     := true,
   directory_   IN VARCHAR2    := null,
   filename_    IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null,
   useXf_       IN BOOLEAN     := false,
   hdr_font_    IN PLS_INTEGER := null,
   hdr_fill_    IN PLS_INTEGER := null,
   col_fmts_    IN numFmt_cols := numFmt_cols() )
IS
   binds_ bind_arr := bind_arr();
BEGIN
   Query2SheetAndAutofilter (
      sql_, binds_, col_headers_, directory_,
      filename_, sheet_, useXf_, hdr_font_, hdr_fill_, col_fmts_
   );
END Query2SheetAndAutofilter;


PROCEDURE SetUseXf (
   p_val BOOLEAN := true )
IS BEGIN
   g_useXf_ := p_val;
END SetUseXf;

------------------------------------------------------------------------------
-- Special Page Generators
-- This must include some font, fill and other initiators too
--

PROCEDURE Init_Workbook
IS
   --gbp_curr_fmt_ VARCHAR2(200) := '_-�* #,##0_-;-�* #,##0_-;_-�* &quot;-&quot;_-;_-@_-';
   gbp_curr_fmt0_ VARCHAR2(200) := '_-&#163;* #,##0_-;-&#163;* #,##0_-;_-&#163;* &quot;-&quot;_-;_-@_-';
   gbp_curr_fmt2_ VARCHAR2(200) := '_-&#163;* #,##0.00_-;-&#163;* #,##0.00_-;_-&#163;* &quot;-&quot;_-;_-@_-';
BEGIN

   Clear_Workbook;
   New_Sheet ('Sheet 1');

   fonts_('head1')       := Get_Font (rgb_ => 'FFDBE5F1', bold_ => true);
   fonts_('bold')        := Get_Font (bold_ => true);
   fonts_('bld_wht')     := Get_Font (rgb_ => 'FFFFFFFF', bold_ => true);
   fonts_('bld_dk_bl')   := Get_Font (rgb_ => 'FF244062', bold_ => true);
   fonts_('bld_lt_bl')   := Get_Font (rgb_ => 'FFDCE6F1', bold_ => true);
   fonts_('bld_ltbl_lg') := Get_Font (rgb_ => 'FFDCE6F1', bold_ => true, fontsize_ => 14);
   fonts_('bld_lt_gr')   := Get_Font (rgb_ => 'FFEBF1DE', bold_ => true);
   fonts_('dk_gr')       := Get_Font (rgb_ => 'FF4F6228');

   fills_('dk_blue')     := Get_Fill ('solid', 'FF17375D');
   fills_('md_dk_blue')  := Get_Fill ('solid', 'FF366092');
   fills_('mid_blue')    := Get_Fill ('solid', 'FF95B3D7');
   fills_('dk_red')      := Get_Fill ('solid', 'FF953735');
   fills_('dk_green')    := Get_Fill ('solid', 'FF006400');
   fills_('lt_green')    := Get_Fill ('solid', 'FFD8E4BC');
   fills_('md_dk_gr')    := Get_Fill ('solid', 'FF76933C');
   fills_('pale_blue')   := Get_Fill ('solid', 'FFDCE6F1');
   fills_('dk_purple')   := Get_Fill ('solid', 'FF60497A');
   fills_('lt_grey')     := Get_Fill ('solid', 'FFD9D9D9');
   fills_('md_grey')     := Get_Fill ('solid', 'FFA6A6A6');
   fills_('dk_grey')     := Get_Fill ('solid', 'FF595959');

   bdrs_('thin')         := Get_Border();
   bdrs_('none')         := Get_Border ('none', 'none', 'none', 'none');
   bdrs_('medium')       := Get_Border ('medium', 'medium', 'medium', 'medium');
   bdrs_('t_medium')     := Get_Border ('medium', 'none', 'none', 'none'); -- top, bottom, left, right
   bdrs_('tl_medium')    := Get_Border ('medium', 'none', 'medium', 'none');
   bdrs_('tr_medium')    := Get_Border ('medium', 'none', 'none', 'medium');
   bdrs_('tb_medium')    := Get_Border ('medium', 'medium', 'none', 'none');
   bdrs_('b_medium')     := Get_Border ('none', 'medium', 'none', 'none');
   bdrs_('bl_medium')    := Get_Border ('none', 'medium', 'medium', 'none');
   bdrs_('l_medium')     := Get_Border ('none', 'none', 'medium', 'none');
   bdrs_('br_medium')    := Get_Border ('none', 'medium', 'none', 'medium');
   bdrs_('r_medium')     := Get_Border ('none', 'none', 'none', 'medium');
   bdrs_('thick')        := Get_Border ('thick', 'thick', 'thick', 'thick');
   bdrs_('t_thick')      := Get_Border ('thick', 'none', 'none', 'none'); -- top, bottom, left, right
   bdrs_('tl_thick')     := Get_Border ('thick', 'none', 'thick', 'none');
   bdrs_('tr_thick')     := Get_Border ('thick', 'none', 'none', 'thick');
   bdrs_('tb_thick')     := Get_Border ('thick', 'thick', 'none', 'none');
   bdrs_('b_thick')      := Get_Border ('none', 'thick', 'none', 'none');
   bdrs_('bl_thick')     := Get_Border ('none', 'thick', 'thick', 'none');
   bdrs_('br_thick')     := Get_Border ('none', 'thick', 'none', 'thick');

   numFmt_('gbp_curr0')  := Get_NumFmt (gbp_curr_fmt0_);
   numFmt_('gbp_curr2')  := Get_NumFmt (gbp_curr_fmt2_);
   numFmt_('2dp')        := Get_NumFmt ('#,##0.00');

   align_('left')        := Get_Alignment (vertical_ => 'center', horizontal_ => 'left',   wrapText_ => false);
   align_('leftw')       := Get_Alignment (vertical_ => 'center', horizontal_ => 'left',   wrapText_ => true);
   align_('right')       := Get_Alignment (vertical_ => 'center', horizontal_ => 'right',  wrapText_ => false);
   align_('center')      := Get_Alignment (vertical_ => 'center', horizontal_ => 'center', wrapText_ => false);
   align_('wrap')        := Get_Alignment (vertical_ => 'top',    horizontal_ => 'left',   wrapText_ => true);

END Init_Workbook;

PROCEDURE Set_Param (
   params_ IN OUT params_arr,
   ix_     IN NUMBER,
   val_    IN VARCHAR2,
   extra_  IN VARCHAR2 := '' )
IS BEGIN
   params_(ix_) := param_rec (
      param_name      => 'Order Number',
      param_value     => val_,
      additional_info => extra_
   );
END Set_Param;

PROCEDURE Bind_Value (
   binds_   IN OUT bind_arr,
   bind_id_ IN VARCHAR2,
   val_     IN VARCHAR2 )
IS BEGIN
   binds_(bind_id_) := data_binder (
      datatype => 'STRING',
      s_val    => val_,
      n_val    => null,
      d_val    => null
   );
END Bind_Value;

PROCEDURE Bind_Value (
   binds_   IN OUT bind_arr,
   bind_id_ IN VARCHAR2,
   val_     IN NUMBER )
IS BEGIN
   binds_(bind_id_) := data_binder (
      datatype => 'NUMBER',
      s_val    => '',
      n_val    => val_,
      d_val    => null
   );
END Bind_Value;

PROCEDURE Bind_Value (
   binds_   IN OUT bind_arr,
   bind_id_ IN VARCHAR2,
   val_     IN DATE )
IS BEGIN
   binds_(bind_id_) := data_binder (
      datatype => 'DATE',
      s_val    => '',
      n_val    => null,
      d_val    => val_
   );
END Bind_Value;

PROCEDURE Create_Params_Sheet (
   report_name_ IN VARCHAR2,
   params_      IN params_arr,
   show_user_   IN BOOLEAN     := true,
   sheet_       IN PLS_INTEGER := null )
IS
   row_ NUMBER := 2;
   sh_  PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN

   -- Information about the report is static, with the only option being as to
   -- whether we show the user who printed the report
   Cell (2, row_, 'Report Information', fontId_ => fonts_('head1'), fillId_ => fills_('dk_blue'), sheet_ => sh_);
   Cell (3, row_, '', fillId_ => fills_('dk_blue'), sheet_ => sh_);
   row_ := row_ + 1;
   Cell (2, row_, 'Report Name', fontId_ => fonts_('bold'), sheet_ => sh_);
   Cell (3, row_, report_name_);
   row_ := row_ + 1;
   Cell (2, row_, 'Executed at', fontId_ => fonts_('bold'), sheet_ => sh_);
   Cell (3, row_, to_char(sysdate, 'YYYY-MM-DD HH24:MI:SS'), sheet_ => sh_);
   row_ := row_ + 1;
   IF show_user_ THEN
      Cell (2, row_, 'Executed by', fontId_ => fonts_('bold'), sheet_ => sh_);
      Cell (3, row_, Fnd_User_API.Get_Description(Fnd_Session_API.Get_Fnd_User), sheet_ => sh_);
      row_ := row_ + 1;
   END IF;

   -- Then we print the parameter headers, with the values output in a loop
   row_ := row_ + 1;
   Cell (2, row_, 'Parameters', fontId_ => fonts_('head1'), fillId_ => fills_('dk_blue'), sheet_ => sh_);
   Cell (3, row_, 'Value', fontId_ => fonts_('head1'), fillId_ => fills_('dk_blue'), sheet_ => sh_);
   Cell (4, row_, 'Additional Info', fontId_ => fonts_('head1'), fillId_ => fills_('dk_blue'), sheet_ => sh_);
   row_ := row_ + 1;
   FOR i_ IN params_.FIRST .. params_.LAST LOOP
      Cell (2, row_, params_(i_).param_name, fontId_ => fonts_('bold'), sheet_ => sh_);
      Cell (3, row_, params_(i_).param_value, sheet_ => sh_);
      Cell (4, row_, params_(i_).additional_info, sheet_ => sh_);
      row_ := row_ + 1;
   END LOOP;

   Set_Column_Width (2, 25, 1);
   Set_Column_Width (3, 40, 1);
   Set_Column_Width (4, 40, 1);

END Create_Params_Sheet;

END AS_XLSX;
/
