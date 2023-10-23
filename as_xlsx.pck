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

TYPE tp_alignment IS RECORD (
   vertical   VARCHAR2(11),
   horizontal VARCHAR2(16),
   wrapText   BOOLEAN );

PROCEDURE Clear_Workbook;

PROCEDURE New_Sheet (
   p_sheetname VARCHAR2 := null,
   p_tabcolor  VARCHAR2 := null );

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 );

FUNCTION OraFmt2Excel (
   p_format IN VARCHAR2 := null ) RETURN VARCHAR2;

FUNCTION get_numFmt (
   p_format IN VARCHAR2 := null ) RETURN PLS_INTEGER;

PROCEDURE Set_Font (
   p_name      IN VARCHAR2,
   p_sheet     IN PLS_INTEGER := null,
   p_family    IN PLS_INTEGER := 2,
   p_fontsize  IN NUMBER := 11,
   p_theme     IN PLS_INTEGER := 1,
   p_underline IN BOOLEAN := false,
   p_italic    IN BOOLEAN := false,
   p_bold      IN BOOLEAN := false,
   p_rgb       IN VARCHAR2 := null ); -- hex Alpha-rgb value

FUNCTION Get_Font (
   p_name      IN VARCHAR2,
   p_family    IN PLS_INTEGER := 2,
   p_fontsize  IN NUMBER      := 11,
   p_theme     IN PLS_INTEGER := 1,
   p_underline IN BOOLEAN     := false,
   p_italic    IN BOOLEAN     := false,
   p_bold      IN BOOLEAN     := false,
   p_rgb       IN VARCHAR2    := null ) RETURN PLS_INTEGER; -- hex Alpha-rgb value

FUNCTION Get_Fill (
   p_patternType IN VARCHAR2,
   p_fgRGB       IN VARCHAR2 := null,                      -- hex Alpha-rgb value
   p_bgRGB       IN VARCHAR2 := null ) RETURN PLS_INTEGER; -- hex Alpha-rgb value

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
   p_vertical   IN VARCHAR2 := null,
   p_horizontal IN VARCHAR2 := null,
   p_wrapText   IN BOOLEAN  := null ) RETURN tp_alignment;

PROCEDURE Cell (
   p_col       IN PLS_INTEGER,
   p_row       IN PLS_INTEGER,
   p_value     IN NUMBER,
   p_numFmtId  IN PLS_INTEGER  := null,
   p_fontId    IN PLS_INTEGER  := null,
   p_fillId    IN PLS_INTEGER  := null,
   p_borderId  IN PLS_INTEGER  := null,
   p_alignment IN tp_alignment := null,
   p_sheet     IN PLS_INTEGER := null );

PROCEDURE Cell (
   p_col       IN PLS_INTEGER,
   p_row       IN PLS_INTEGER,
   p_value     IN VARCHAR2,
   p_numFmtId  IN PLS_INTEGER  := null,
   p_fontId    IN PLS_INTEGER  := null,
   p_fillId    IN PLS_INTEGER  := null,
   p_borderId  IN PLS_INTEGER  := null,
   p_alignment IN tp_alignment := null,
   p_sheet     IN PLS_INTEGER  := null );

PROCEDURE Cell (
   p_col       IN PLS_INTEGER,
   p_row       IN PLS_INTEGER,
   p_value     IN DATE,
   p_numFmtId  IN PLS_INTEGER  := null,
   p_fontId    IN PLS_INTEGER  := null,
   p_fillId    IN PLS_INTEGER  := null,
   p_borderId  IN PLS_INTEGER  := null,
   p_alignment IN tp_alignment := null,
   p_sheet     IN PLS_INTEGER := null );

PROCEDURE Hyperlink (
   p_col   IN PLS_INTEGER,
   p_row   IN PLS_INTEGER,
   p_url   IN VARCHAR2,
   p_value IN VARCHAR2    := null,
   p_sheet IN PLS_INTEGER := null );

PROCEDURE Comment (
   p_col    IN PLS_INTEGER,
   p_row    IN PLS_INTEGER,
   p_text   IN VARCHAR2,
   p_author IN VARCHAR2 := null,
   p_width  IN PLS_INTEGER := 150,  -- pixels
   p_height IN PLS_INTEGER := 100,  -- pixels
   p_sheet  IN PLS_INTEGER := null );

PROCEDURE Num_Formula (
   p_col           IN PLS_INTEGER,
   p_row           IN PLS_INTEGER,
   p_formula       IN VARCHAR2,
   p_default_value IN NUMBER       := null,
   p_numFmtId      IN PLS_INTEGER  := null,
   p_fontId        IN PLS_INTEGER  := null,
   p_fillId        IN PLS_INTEGER  := null,
   p_borderId      IN PLS_INTEGER  := null,
   p_alignment     IN tp_alignment := null,
   p_sheet         IN PLS_INTEGER  := null );

PROCEDURE Str_Formula (
   p_col           IN PLS_INTEGER,
   p_row           IN PLS_INTEGER,
   p_formula       IN VARCHAR2,
   p_default_value IN VARCHAR2     := null,
   p_numFmtId      IN PLS_INTEGER  := null,
   p_fontId        IN PLS_INTEGER  := null,
   p_fillId        IN PLS_INTEGER  := null,
   p_borderId      IN PLS_INTEGER  := null,
   p_alignment     IN tp_alignment := null,
   p_sheet         IN PLS_INTEGER := null );

PROCEDURE Mergecells (
   p_tl_col  IN PLS_INTEGER, -- top left
   p_tl_row  IN PLS_INTEGER,
   p_br_col  IN PLS_INTEGER, -- bottom right
   p_br_row  IN PLS_INTEGER,
   p_sheet   IN PLS_INTEGER := null );

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
   p_sheet        IN PLS_INTEGER := null );

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
   p_sheet        IN PLS_INTEGER := null );

PROCEDURE Defined_Name (
   p_tl_col     IN PLS_INTEGER, -- top left
   p_tl_row     IN PLS_INTEGER,
   p_br_col     IN PLS_INTEGER, -- bottom right
   p_br_row     IN PLS_INTEGER,
   p_name       IN VARCHAR2,
   p_sheet      IN PLS_INTEGER := null,
   p_localsheet IN PLS_INTEGER := null );

PROCEDURE Set_Column_Width (
   p_col   IN PLS_INTEGER,
   p_width IN NUMBER,
   p_sheet IN PLS_INTEGER := null );

PROCEDURE Set_Column (
   p_col       IN PLS_INTEGER,
   p_numFmtId  IN PLS_INTEGER  := null,
   p_fontId    IN PLS_INTEGER  := null,
   p_fillId    IN PLS_INTEGER  := null,
   p_borderId  IN PLS_INTEGER  := null,
   p_alignment IN tp_alignment := null,
   p_sheet     IN PLS_INTEGER  := null );

PROCEDURE Set_Row (
   p_row       IN PLS_INTEGER,
   p_numFmtId  IN PLS_INTEGER  := null,
   p_fontId    IN PLS_INTEGER  := null,
   p_fillId    IN PLS_INTEGER  := null,
   p_borderId  IN PLS_INTEGER  := null,
   p_alignment IN tp_alignment := null,
   p_sheet     IN PLS_INTEGER  := null,
   p_height    IN NUMBER := null );

PROCEDURE Freeze_Rows (
   p_nr_rows  IN PLS_INTEGER := 1,
   p_sheet    IN PLS_INTEGER := null );

PROCEDURE Freeze_Cols (
   p_nr_cols IN PLS_INTEGER := 1,
   p_sheet   IN PLS_INTEGER := null );

PROCEDURE Freeze_Pane (
   p_col PLS_INTEGER,
   p_row PLS_INTEGER,
   p_sheet PLS_INTEGER := null );

PROCEDURE Set_Autofilter (
   p_column_start PLS_INTEGER := null,
   p_column_end PLS_INTEGER := null,
   p_row_start PLS_INTEGER := null,
   p_row_end PLS_INTEGER := null,
   p_sheet PLS_INTEGER := null );

PROCEDURE Set_Tabcolor (
   p_tabcolor VARCHAR2, -- hex Alpha-rgb value
   p_sheet PLS_INTEGER := null );

FUNCTION Finish RETURN BLOB;

PROCEDURE Save (
   p_directory VARCHAR2,
   p_filename VARCHAR2 );

PROCEDURE query2sheet (
   sql_            IN VARCHAR2,
   column_headers_ IN BOOLEAN     := true,
   directory_      IN VARCHAR2    := null,
   filename_       IN VARCHAR2    := null,
   sheet_          IN PLS_INTEGER := null,
   UseXf_          IN BOOLEAN     := false );

PROCEDURE query2sheet (
   p_rc             IN OUT SYS_REFCURSOR,
   p_column_headers IN BOOLEAN     := true,
   p_directory      IN VARCHAR2    := null,
   p_filename       IN VARCHAR2    := null,
   p_sheet          IN PLS_INTEGER := null,
   p_UseXf          IN BOOLEAN := false );

PROCEDURE setUseXf (
   p_val BOOLEAN := true );

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
g_useXf               BOOLEAN := true;
g_addtxt2utf8blob_tmp VARCHAR2(32767);

PROCEDURE addtxt2utf8blob_init (
   p_blob IN OUT NOCOPY BLOB )
IS BEGIN
   g_addtxt2utf8blob_tmp := null;
   dbms_lob.createtemporary( p_blob, true );
END addtxt2utf8blob_init;

PROCEDURE Addtxt2utf8blob_Finish (
   p_blob IN OUT NOCOPY BLOB )
IS
   t_raw RAW(32767);
BEGIN
   t_raw := utl_i18n.string_to_raw(g_addtxt2utf8blob_tmp, 'AL32UTF8');
   dbms_lob.writeappend (p_blob, utl_raw.length(t_raw), t_raw);
EXCEPTION
   WHEN value_error THEN
      t_raw := utl_i18n.string_to_raw(substr(g_addtxt2utf8blob_tmp,1,16381), 'AL32UTF8');
      dbms_lob.writeappend (p_blob, utl_raw.length(t_raw), t_raw);
      t_raw := utl_i18n.string_to_raw(substr(g_addtxt2utf8blob_tmp,16382), 'AL32UTF8');
      dbms_lob.writeappend (p_blob, utl_raw.length(t_raw), t_raw);
END addtxt2utf8blob_finish;

PROCEDURE addtxt2utf8blob( p_txt VARCHAR2, p_blob in out nocopy blob )
IS BEGIN
   g_addtxt2utf8blob_tmp := g_addtxt2utf8blob_tmp || p_txt;
EXCEPTION
   WHEN value_error THEN
      addtxt2utf8blob_finish(p_blob);
      g_addtxt2utf8blob_tmp := p_txt;
END addtxt2utf8blob;
--
  PROCEDURE blob2file
    ( p_blob blob
    , p_directory VARCHAR2 := 'MY_DIR'
    , p_filename VARCHAR2 := 'my.xlsx'
    )
  is
    t_fh utl_file.file_TYPE;
    t_len PLS_INTEGER := 32767;
  begin
    t_fh := utl_file.fopen( p_directory
                          , p_filename
                          , 'wb'
                          );
    for i in 0 .. trunc( ( dbms_lob.getlength( p_blob ) - 1 ) / t_len )
    loop
      utl_file.put_raw( t_fh
                      , dbms_lob.substr( p_blob
                                       , t_len
                                       , i * t_len + 1
                                       )
                      );
    end loop;
    utl_file.fclose( t_fh );
  end;
--
  function raw2num( p_raw raw, p_len integer, p_pos integer )
  return NUMBER
  is
  begin
    return utl_raw.cast_to_binary_integer( utl_raw.substr( p_raw, p_pos, p_len ), utl_raw.little_endian );
  end;
--
  function little_endian( p_big NUMBER, p_bytes PLS_INTEGER := 4 )
  return raw
  is
  begin
    return utl_raw.substr( utl_raw.cast_from_binary_integer( p_big, utl_raw.little_endian ), 1, p_bytes );
  end;
--
  function blob2num( p_blob blob, p_len integer, p_pos integer )
  return NUMBER
  is
  begin
    return utl_raw.cast_to_binary_integer( dbms_lob.substr( p_blob, p_len, p_pos ), utl_raw.little_endian );
  end;
--
  PROCEDURE add1file
    ( p_zipped_blob in out blob
    , p_name VARCHAR2
    , p_content blob
    )
  is
    t_now date;
    t_blob blob;
    t_len integer;
    t_clen integer;
    t_crc32 raw(4) := hextoraw( '00000000' );
    t_compressed BOOLEAN := false;
    t_name raw(32767);
  begin
    t_now := sysdate;
    t_len := nvl( dbms_lob.getlength( p_content ), 0 );
    if t_len > 0
    then
      t_blob := utl_compress.lz_compress( p_content );
      t_clen := dbms_lob.getlength( t_blob ) - 18;
      t_compressed := t_clen < t_len;
      t_crc32 := dbms_lob.substr( t_blob, 4, t_clen + 11 );
    end if;
    if not t_compressed
    then
      t_clen := t_len;
      t_blob := p_content;
    end if;
    if p_zipped_blob is null
    then
      dbms_lob.createtemporary( p_zipped_blob, true );
    end if;
    t_name := utl_i18n.string_to_raw( p_name, 'AL32UTF8' );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( LOCAL_FILE_HEADER_ -- Local file header signature
                                   , hextoraw( '1400' )  -- version 2.0
                                   , case when t_name = utl_i18n.string_to_raw( p_name, 'US8PC437' )
                                       then hextoraw( '0000' ) -- no General purpose bits
                                       else hextoraw( '0008' ) -- set Language encoding flag (EFS)
                                     end
                                   , case when t_compressed
                                        then hextoraw( '0800' ) -- deflate
                                        else hextoraw( '0000' ) -- stored
                                     end
                                   , little_endian( to_number( to_char( t_now, 'ss' ) ) / 2
                                                  + to_number( to_char( t_now, 'mi' ) ) * 32
                                                  + to_number( to_char( t_now, 'hh24' ) ) * 2048
                                                  , 2
                                                  ) -- File last modification time
                                   , little_endian( to_number( to_char( t_now, 'dd' ) )
                                                  + to_number( to_char( t_now, 'mm' ) ) * 32
                                                  + ( to_number( to_char( t_now, 'yyyy' ) ) - 1980 ) * 512
                                                  , 2
                                                  ) -- File last modification date
                                   , t_crc32 -- CRC-32
                                   , little_endian( t_clen )                      -- compressed size
                                   , little_endian( t_len )                       -- uncompressed size
                                   , little_endian( utl_raw.length( t_name ), 2 ) -- File name length
                                   , hextoraw( '0000' )                           -- Extra field length
                                   , t_name                                       -- File name
                                   )
                   );
    if t_compressed
    then
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 11 ); -- compressed content
    elsif t_clen > 0
    then
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 1 ); --  content
    end if;
    if dbms_lob.istemporary( t_blob ) = 1
    then
      dbms_lob.freetemporary( t_blob );
    end if;
  end;
--
PROCEDURE Finish_Zip (
   p_zipped_blob IN OUT BLOB )
IS
   t_cnt             PLS_INTEGER := 0;
   t_offs            INTEGER;
   t_offs_dir_header INTEGER;
   t_offs_end_header INTEGER;
   t_comment         RAW(200) := Utl_Raw.Cast_To_Raw(
      'Implementation by Anton Scheffer, ' || VERSION_
   );
BEGIN
   t_offs_dir_header := dbms_lob.getlength (p_zipped_blob);
   t_offs := 1;
   WHILE Dbms_Lob.Substr(p_zipped_blob, utl_raw.length(LOCAL_FILE_HEADER_), t_offs) = LOCAL_FILE_HEADER_ LOOP
      t_cnt := t_cnt + 1;
      Dbms_Lob.Append (
         p_zipped_blob,
         Utl_Raw.Concat (
            hextoraw('504B0102'),      -- Central directory file header signature
            hextoraw('1400'),          -- version 2.0
            dbms_lob.substr(p_zipped_blob, 26, t_offs+4),
            hextoraw('0000'),          -- File comment length
            hextoraw('0000'),          -- Disk number where file starts
            hextoraw('0000'),          -- Internal file attributes => 0000=binary-file; 0100(ascii)=text-file
            CASE
               WHEN Dbms_Lob.Substr (
                  p_zipped_blob, 1, t_offs+30+blob2num(p_zipped_blob,2,t_offs+26)-1
               ) IN (hextoraw('2F'), hextoraw('5C'))
               THEN
                  hextoraw('10000000') -- a directory/folder
               ELSE
                  hextoraw('2000B681') -- a file
            END,                       -- External file attributes
            little_endian(t_offs-1),   -- Relative offset of local file header
            dbms_lob.substr(p_zipped_blob, blob2num(p_zipped_blob,2,t_offs+26),t_offs+30) -- File name
         )
      );
      t_offs := t_offs + 30 +
         blob2num(p_zipped_blob, 4, t_offs+18 ) + -- compressed size
         blob2num(p_zipped_blob, 2, t_offs+26 ) + -- File name length
         blob2num(p_zipped_blob, 2, t_offs+28 );  -- Extra field length
   END LOOP;
   t_offs_end_header := dbms_lob.getlength(p_zipped_blob);
   Dbms_Lob.Append (
       p_zipped_blob,
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
   p_tabcolor VARCHAR2,
   p_sheet    PLS_INTEGER := null )
IS
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   workbook.sheets(t_sheet).tabcolor := substr( p_tabcolor, 1, 8 );
END Set_Tabcolor;

PROCEDURE New_Sheet (
   p_sheetname VARCHAR2 := null,
   p_tabcolor  VARCHAR2 := null )
IS
   s_  PLS_INTEGER := workbook.sheets.count() + 1;
BEGIN
   workbook.sheets(s_).name := nvl(dbms_xmlgen.convert(translate(p_sheetname, 'a/\[]*:?', 'a')), 'Sheet'||s_);
   IF workbook.strings.count() = 0 THEN
      workbook.str_cnt := 0;
   END IF;
   IF workbook.fonts.count() = 0 THEN
      workbook.fontid := get_font('Calibri');
   END IF;
   IF workbook.fills.count() = 0 THEN
      Get_Fill('none');
      Get_Fill('gray125');
   END IF;
   IF workbook.borders.count() = 0 THEN
      Get_Border ('', '', '', '');
   END IF;
   set_tabcolor(p_tabcolor, s_);
   workbook.sheets(s_).fontid := workbook.fontid;
END New_Sheet;

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 )
IS BEGIN
   workbook.sheets(sheet_).name := nvl(dbms_xmlgen.convert(translate(name_, 'a/\[]*:?', 'a')), 'Sheet'||sheet_);
END Set_Sheet_Name;

PROCEDURE Set_Col_Width (
   p_sheet  PLS_INTEGER,
   p_col    PLS_INTEGER,
   p_format VARCHAR2 )
IS
   t_width  NUMBER;
   t_nr_chr PLS_INTEGER;
BEGIN
   IF p_format IS null THEN
      RETURN;
   END IF;
   IF instr(p_format, ';') > 0 THEN
      t_nr_chr := length( translate( substr( p_format, 1, instr( p_format, ';' ) - 1 ), 'a\"', 'a' ) );
   ELSE
      t_nr_chr := length( translate( p_format, 'a\"', 'a' ) );
   END IF;
   t_width := trunc((t_nr_chr * 7 + 5 ) / 7 * 256 ) / 256; -- assume default 11 point Calibri
   IF workbook.sheets(p_sheet).widths.exists(p_col) THEN
      workbook.sheets(p_sheet).widths(p_col) := greatest(
         workbook.sheets(p_sheet).widths(p_col), t_width
      );
   ELSE
      workbook.sheets(p_sheet).widths(p_col) := greatest(t_width, 8.43);
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

FUNCTION get_numFmt (
   p_format VARCHAR2 := null) RETURN PLS_INTEGER
IS
   t_cnt      PLS_INTEGER;
   t_numFmtId PLS_INTEGER;
BEGIN
   IF p_format is null THEN
      RETURN 0;
   END IF;
   t_cnt := workbook.numFmts.count();
   FOR i_ in 1 .. t_cnt LOOP
      IF workbook.numFmts(i_).formatCode = p_format THEN
        t_numFmtId := workbook.numFmts(i_).numFmtId;
        EXIT;
      END IF;
   END LOOP;
   IF t_numFmtId is null THEN
      t_numFmtId := CASE WHEN t_cnt = 0 THEN 164 ELSE workbook.numFmts( t_cnt ).numFmtId + 1 END;
      t_cnt := t_cnt + 1;
      workbook.numFmts(t_cnt).numFmtId     := t_numFmtId;
      workbook.numFmts(t_cnt).formatCode   := p_format;
      workbook.numFmtIndexes( t_numFmtId ) := t_cnt;
   END IF;
   RETURN t_numFmtId;
END get_numFmt;

PROCEDURE Set_Font (
   p_name      VARCHAR2,
   p_sheet     PLS_INTEGER := null,
   p_family    PLS_INTEGER := 2,
   p_fontsize  NUMBER := 11,
   p_theme     PLS_INTEGER := 1,
   p_underline BOOLEAN := false,
   p_italic    BOOLEAN := false,
   p_bold      BOOLEAN := false,
   p_rgb       VARCHAR2 := null ) -- this is a hex ALPHA Red Green Blue value
IS
   t_ind PLS_INTEGER := get_font (p_name, p_family, p_fontsize, p_theme, p_underline, p_italic, p_bold, p_rgb);
BEGIN
   IF p_sheet IS null THEN
      workbook.fontid := t_ind;
   ELSE
      workbook.sheets( p_sheet ).fontid := t_ind;
   END IF;
END Set_Font;

FUNCTION Get_Font (
   p_name      VARCHAR2,
   p_family    PLS_INTEGER := 2,
   p_fontsize  NUMBER := 11,
   p_theme     PLS_INTEGER := 1,
   p_underline BOOLEAN := false,
   p_italic    BOOLEAN := false,
   p_bold      BOOLEAN := false,
   p_rgb       VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   t_ind PLS_INTEGER;
BEGIN
   IF workbook.fonts.count() > 0 THEN
      FOR f IN 0 .. workbook.fonts.count() - 1 LOOP
         IF (     workbook.fonts(f).name = p_name
              AND workbook.fonts(f).family = p_family
              AND workbook.fonts(f).fontsize = p_fontsize
              AND workbook.fonts(f).theme = p_theme
              AND workbook.fonts(f).underline = p_underline
              AND workbook.fonts(f).italic = p_italic
              AND workbook.fonts(f).bold = p_bold
              AND (      workbook.fonts(f).rgb = p_rgb
                    OR ( workbook.fonts(f).rgb IS null AND p_rgb IS null )
              )
         ) THEN
            RETURN f;
         END IF;
      END LOOP;
   END IF;
   t_ind := workbook.fonts.count();
   workbook.fonts(t_ind).name      := p_name;
   workbook.fonts(t_ind).family    := p_family;
   workbook.fonts(t_ind).fontsize  := p_fontsize;
   workbook.fonts(t_ind).theme     := p_theme;
   workbook.fonts(t_ind).underline := p_underline;
   workbook.fonts(t_ind).italic    := p_italic;
   workbook.fonts(t_ind).bold      := p_bold;
   workbook.fonts(t_ind).rgb       := p_rgb;
   RETURN t_ind;
END Get_Font;


FUNCTION Get_Fill (
   p_patternType VARCHAR2,
   p_fgRGB       VARCHAR2 := null,
   p_bgRGB       VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   t_ind PLS_INTEGER;
BEGIN
   IF workbook.fills.count() > 0 THEN
      FOR f IN 0 .. workbook.fills.count() - 1 LOOP
         IF (   workbook.fills(f).patternType = p_patternType
            AND nvl(workbook.fills(f).fgRGB, 'x') = nvl(upper(p_fgRGB), 'x')
            AND nvl(workbook.fills(f).bgRGB, 'x') = nvl(upper(p_bgRGB), 'x')
         ) THEN
            RETURN f;
         END IF;
      END LOOP;
   END IF;
   t_ind := workbook.fills.count();
   workbook.fills(t_ind).patternType := p_patternType;
   workbook.fills(t_ind).fgRGB := upper(p_fgRGB);
   workbook.fills(t_ind).bgRGB := upper(p_bgRGB);
   RETURN t_ind;
END Get_Fill;

PROCEDURE Get_Fill (
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,
   bgRGB_       IN VARCHAR2 := null )
IS
   ix_ PLS_INTEGER := Get_Fill (patternType_, fgRGB_, bgRGB_);
BEGIN
   Dbms_Output.Put_Line (ix_);
END Get_Fill;


FUNCTION Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' ) RETURN PLS_INTEGER
IS
   t_ind PLS_INTEGER;
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
   t_ind := workbook.borders.count();
   workbook.borders(t_ind).top    := top_;
   workbook.borders(t_ind).bottom := bottom_;
   workbook.borders(t_ind).left   := left_;
   workbook.borders(t_ind).right  := right_;
   RETURN t_ind;
END Get_Border;

PROCEDURE Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' )
IS
   ix_ NUMBER := Get_Border (top_, bottom_, left_, right_);
BEGIN
   Dbms_Output.Put_Line (ix_); -- avoid compiler warning
END Get_Border;

FUNCTION Get_Alignment (
   p_vertical    VARCHAR2 := null,
   p_horizontal  VARCHAR2 := null,
   p_wrapText    BOOLEAN := null ) RETURN tp_alignment
IS
   t_rv tp_alignment;
BEGIN
   t_rv.vertical := p_vertical;
   t_rv.horizontal := p_horizontal;
   t_rv.wrapText := p_wrapText;
   RETURN t_rv;
END Get_Alignment;

FUNCTION get_XfId (
   p_sheet     PLS_INTEGER,
   p_col       PLS_INTEGER,
   p_row       PLS_INTEGER,
   p_numFmtId  PLS_INTEGER := null,
   p_fontId    PLS_INTEGER := null,
   p_fillId    PLS_INTEGER := null,
   p_borderId  PLS_INTEGER := null,
   p_alignment tp_alignment := null ) RETURN VARCHAR2
IS
   t_cnt    PLS_INTEGER;
   t_XfId   PLS_INTEGER;
   t_XF     tp_XF_fmt;
   t_col_XF tp_XF_fmt;
   t_row_XF tp_XF_fmt;
BEGIN
   IF not g_useXf THEN
      RETURN '';
   END IF;
   IF workbook.sheets(p_sheet).col_fmts.exists(p_col) THEN
      t_col_XF := workbook.sheets(p_sheet).col_fmts(p_col);
   END IF;
   IF workbook.sheets(p_sheet).row_fmts.exists(p_row) THEN
      t_row_XF := workbook.sheets( p_sheet ).row_fmts( p_row );
   END IF;
   t_XF.numFmtId := coalesce (p_numFmtId, t_col_XF.numFmtId, t_row_XF.numFmtId, workbook.sheets(p_sheet).fontid, workbook.fontid);
   t_XF.fontId   := coalesce (p_fontId, t_col_XF.fontId, t_row_XF.fontId, 0);
   t_XF.fillId   := coalesce (p_fillId, t_col_XF.fillId, t_row_XF.fillId, 0);
   t_XF.borderId := coalesce (p_borderId, t_col_XF.borderId, t_row_XF.borderId, 0);
   t_XF.alignment := Get_Alignment (
      coalesce( p_alignment.vertical, t_col_XF.alignment.vertical, t_row_XF.alignment.vertical ),
      coalesce( p_alignment.horizontal, t_col_XF.alignment.horizontal, t_row_XF.alignment.horizontal ),
      coalesce( p_alignment.wrapText, t_col_XF.alignment.wrapText, t_row_XF.alignment.wrapText )
   );
   IF t_XF.numFmtId + t_XF.fontId + t_XF.fillId + t_XF.borderId = 0
      AND t_XF.alignment.vertical IS null
      AND t_XF.alignment.horizontal IS null
      AND not nvl(t_XF.alignment.wrapText, false)
   THEN
      RETURN '';
   END IF;
   IF t_XF.numFmtId > 0 THEN
      set_col_width (p_sheet, p_col, workbook.numFmts(workbook.numFmtIndexes(t_XF.numFmtId)).formatCode);
   END IF;
   t_cnt := workbook.cellXfs.count();
   FOR i IN 1 .. t_cnt LOOP
      IF (   workbook.cellXfs(i).numFmtId = t_XF.numFmtId
         and workbook.cellXfs(i).fontId = t_XF.fontId
         and workbook.cellXfs(i).fillId = t_XF.fillId
         and workbook.cellXfs(i).borderId = t_XF.borderId
         and nvl( workbook.cellXfs(i).alignment.vertical, 'x') = nvl( t_XF.alignment.vertical, 'x')
         and nvl( workbook.cellXfs(i).alignment.horizontal, 'x') = nvl( t_XF.alignment.horizontal, 'x')
         and nvl( workbook.cellXfs(i).alignment.wrapText, false) = nvl( t_XF.alignment.wrapText, false)
      ) THEN
         t_XfId := i;
         exit;
      END IF;
   END LOOP;
   IF t_XfId IS null THEN
      t_cnt := t_cnt + 1;
      t_XfId := t_cnt;
      workbook.cellXfs( t_cnt ) := t_XF;
   END IF;
   RETURN 's="' || t_XfId || '"';
END get_XfId;
--
PROCEDURE Cell (
   p_col PLS_INTEGER,
   p_row PLS_INTEGER,
   p_value NUMBER,
   p_numFmtId PLS_INTEGER := null,
   p_fontId PLS_INTEGER := null,
   p_fillId PLS_INTEGER := null,
   p_borderId PLS_INTEGER := null,
   p_alignment tp_alignment := null,
   p_sheet PLS_INTEGER := null )
IS
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   workbook.sheets(t_sheet).rows(p_row)(p_col).value := p_value;
   workbook.sheets(t_sheet).rows(p_row)(p_col).style := null;
   workbook.sheets(t_sheet).rows(p_row)(p_col).style := get_XfId (
      t_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment
   );
END Cell;
--
FUNCTION Add_String (
   p_string VARCHAR2 ) RETURN PLS_INTEGER
IS
   t_cnt PLS_INTEGER;
BEGIN
   IF workbook.strings.exists(nvl( p_string, '')) THEN
      t_cnt := workbook.strings(nvl( p_string, ''));
   ELSE
      t_cnt := workbook.strings.count();
      workbook.str_ind(t_cnt) := p_string;
      workbook.strings(nvl(p_string, '')) := t_cnt;
   END IF;
   workbook.str_cnt := workbook.str_cnt + 1;
   RETURN t_cnt;
END Add_String;

PROCEDURE Cell (
   p_col       PLS_INTEGER,
   p_row       PLS_INTEGER,
   p_value     VARCHAR2,
   p_numFmtId  PLS_INTEGER  := null,
   p_fontId    PLS_INTEGER  := null,
   p_fillId    PLS_INTEGER  := null,
   p_borderId  PLS_INTEGER  := null,
   p_alignment tp_alignment := null,
   p_sheet     PLS_INTEGER  := null )
IS
   t_sheet     PLS_INTEGER  := nvl(p_sheet, workbook.sheets.count());
   t_alignment tp_alignment := p_alignment;
BEGIN
   workbook.sheets(t_sheet).rows(p_row)(p_col).value := add_string(p_value);
   IF t_alignment.wrapText IS null AND instr( p_value, chr(13) ) > 0 THEN
      t_alignment.wrapText := true;
   END IF;
   workbook.sheets(t_sheet).rows(p_row)(p_col).style := 't="s" ' || get_XfId (
      t_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, t_alignment
   );
END Cell;

PROCEDURE Cell (
   p_col       PLS_INTEGER,
   p_row       PLS_INTEGER,
   p_value     DATE,
   p_numFmtId  PLS_INTEGER  := null,
   p_fontId    PLS_INTEGER  := null,
   p_fillId    PLS_INTEGER  := null,
   p_borderId  PLS_INTEGER  := null,
   p_alignment tp_alignment := null,
   p_sheet     PLS_INTEGER  := null )
IS
   t_numFmtId PLS_INTEGER := p_numFmtId;
   t_sheet    PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   workbook.sheets(t_sheet).rows(p_row)(p_col).value := (p_value - date '1900-03-01' ) + 61;
   IF t_numFmtId IS null
      AND not (    workbook.sheets(t_sheet).col_fmts.exists(p_col)
               AND workbook.sheets(t_sheet).col_fmts(p_col).numFmtId IS not null )
      AND not (    workbook.sheets(t_sheet).row_fmts.exists(p_row)
               and workbook.sheets(t_sheet).row_fmts(p_row).numFmtId IS not null )
   THEN
      t_numFmtId := get_numFmt('dd/mm/yyyy');
   END IF;
   workbook.sheets(t_sheet).rows(p_row)(p_col).style := get_XfId (
      t_sheet, p_col, p_row, t_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment
   );
END Cell;

PROCEDURE Query_Date_Cell (
   p_col   PLS_INTEGER,
   p_row   PLS_INTEGER,
   p_value DATE,
   p_sheet PLS_INTEGER := null,
   p_XfId VARCHAR2 )
IS
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   Cell (p_col, p_row, p_value, 0, p_sheet => t_sheet);
   workbook.sheets(t_sheet).rows(p_row)(p_col).style := p_XfId;
END Query_Date_Cell;

PROCEDURE Hyperlink (
   p_col   PLS_INTEGER,
   p_row   PLS_INTEGER,
   p_url   VARCHAR2,
   p_value VARCHAR2    := null,
   p_sheet PLS_INTEGER := null )
IS
   t_ind   PLS_INTEGER;
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   workbook.sheets(t_sheet).rows(p_row)(p_col).value := add_string(nvl(p_value, p_url));
   workbook.sheets(t_sheet).rows(p_row)(p_col).style := 't="s" ' || get_XfId( t_sheet, p_col, p_row, '', Get_Font('Calibri', p_theme => 10, p_underline => true));
   t_ind := workbook.sheets(t_sheet).hyperlinks.count() + 1;
   workbook.sheets(t_sheet).hyperlinks(t_ind).cell := alfan_col(p_col) || p_row;
   workbook.sheets(t_sheet).hyperlinks(t_ind).url := p_url;
END Hyperlink;

PROCEDURE Comment (
   p_col    PLS_INTEGER,
   p_row    PLS_INTEGER,
   p_text   VARCHAR2,
   p_author VARCHAR2 := null,
   p_width  PLS_INTEGER := 150,
   p_height PLS_INTEGER := 100,
   p_sheet PLS_INTEGER := null )
IS
   t_ind   PLS_INTEGER;
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   t_ind := workbook.sheets(t_sheet).comments.count() + 1;
   workbook.sheets(t_sheet).comments(t_ind).row    := p_row;
   workbook.sheets(t_sheet).comments(t_ind).column := p_col;
   workbook.sheets(t_sheet).comments(t_ind).text   := dbms_xmlgen.convert(p_text);
   workbook.sheets(t_sheet).comments(t_ind).author := dbms_xmlgen.convert(p_author);
   workbook.sheets(t_sheet).comments(t_ind).width  := p_width;
   workbook.sheets(t_sheet).comments(t_ind).height := p_height;
END Comment;

  PROCEDURE num_formula
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_formula VARCHAR2
    , p_default_value NUMBER := null
    , p_numFmtId PLS_INTEGER := null
    , p_fontId PLS_INTEGER := null
    , p_fillId PLS_INTEGER := null
    , p_borderId PLS_INTEGER := null
    , p_alignment tp_alignment := null
    , p_sheet PLS_INTEGER := null
    )
  is
    t_ind PLS_INTEGER := workbook.formulas.count;
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  begin
    workbook.formulas( t_ind ) := p_formula;
    cell( p_col, p_row, p_default_value, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment, t_sheet );
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).formula_idx := t_ind;
  end;
--
  PROCEDURE str_formula
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_formula VARCHAR2
    , p_default_value VARCHAR2 := null
    , p_numFmtId PLS_INTEGER := null
    , p_fontId PLS_INTEGER := null
    , p_fillId PLS_INTEGER := null
    , p_borderId PLS_INTEGER := null
    , p_alignment tp_alignment := null
    , p_sheet PLS_INTEGER := null
    )
  is
    t_ind PLS_INTEGER := workbook.formulas.count;
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  begin
    workbook.formulas( t_ind ) := p_formula;
    cell( p_col, p_row, p_default_value, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment, t_sheet );
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).formula_idx := t_ind;
  end;

PROCEDURE Mergecells (
   p_tl_col IN PLS_INTEGER, -- top left
   p_tl_row IN PLS_INTEGER,
   p_br_col IN PLS_INTEGER, -- bottom right
   p_br_row IN PLS_INTEGER,
   p_sheet  IN PLS_INTEGER := null )
IS
   t_ind   PLS_INTEGER;
   t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
BEGIN
    t_ind := workbook.sheets( t_sheet ).mergecells.count() + 1;
    workbook.sheets( t_sheet ).mergecells( t_ind ) := alfan_col( p_tl_col ) || p_tl_row || ':' || alfan_col( p_br_col ) || p_br_row;
END Mergecells;
--
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
   p_sheet       IN PLS_INTEGER := null )
IS
   ix_     PLS_INTEGER;
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   ix_ := workbook.sheets(t_sheet).validations.count() + 1;
   workbook.sheets(t_sheet).validations(ix_).type        := p_type;
   workbook.sheets(t_sheet).validations(ix_).errorstyle  := p_style;
   workbook.sheets(t_sheet).validations(ix_).sqref       := p_sqref;
   workbook.sheets(t_sheet).validations(ix_).formula1    := p_formula1;
   workbook.sheets(t_sheet).validations(ix_).formula2    := p_formula2;
   workbook.sheets(t_sheet).validations(ix_).error_title := p_error_title;
   workbook.sheets(t_sheet).validations(ix_).error_txt   := p_error_txt;
   workbook.sheets(t_sheet).validations(ix_).title       := p_title;
   workbook.sheets(t_sheet).validations(ix_).prompt      := p_prompt;
   workbook.sheets(t_sheet).validations(ix_).showerrormessage := p_show_error;
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
   p_sheet        IN PLS_INTEGER := null )
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
      p_sheet       => p_sheet
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
   p_sheet        IN PLS_INTEGER := null )
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
      p_sheet       => p_sheet
   );
END List_Validation;
--
PROCEDURE Defined_Name (
   p_tl_col     PLS_INTEGER, -- top left
   p_tl_row     PLS_INTEGER,
   p_br_col     PLS_INTEGER, -- bottom right
   p_br_row     PLS_INTEGER,
   p_name       VARCHAR2,
   p_sheet      PLS_INTEGER := null,
   p_localsheet PLS_INTEGER := null )
IS
   t_ind   PLS_INTEGER;
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   t_ind := workbook.defined_names.count() + 1;
   workbook.defined_names(t_ind).name := p_name;
   workbook.defined_names(t_ind).ref := 'Sheet' || t_sheet || '!$' || alfan_col( p_tl_col ) || '$' ||  p_tl_row || ':$' || alfan_col( p_br_col ) || '$' || p_br_row;
   workbook.defined_names(t_ind).sheet := p_localsheet;
END Defined_Name;

PROCEDURE Set_Column_Width (
   p_col   PLS_INTEGER,
   p_width NUMBER,
   p_sheet PLS_INTEGER := null )
IS
   t_width NUMBER;
BEGIN
   t_width := trunc( round( p_width * 7 ) * 256 / 7 ) / 256;
   workbook.sheets( nvl( p_sheet, workbook.sheets.count() ) ).widths( p_col ) := t_width;
END Set_Column_Width;

PROCEDURE Set_Column (
   p_col       PLS_INTEGER,
   p_numFmtId  PLS_INTEGER  := null,
   p_fontId    PLS_INTEGER  := null,
   p_fillId    PLS_INTEGER  := null,
   p_borderId  PLS_INTEGER  := null,
   p_alignment tp_alignment := null,
   p_sheet     PLS_INTEGER  := null )
IS
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   workbook.sheets(t_sheet).col_fmts(p_col).numFmtId  := p_numFmtId;
   workbook.sheets(t_sheet).col_fmts(p_col).fontId    := p_fontId;
   workbook.sheets(t_sheet).col_fmts(p_col).fillId    := p_fillId;
   workbook.sheets(t_sheet).col_fmts(p_col).borderId  := p_borderId;
   workbook.sheets(t_sheet).col_fmts(p_col).alignment := p_alignment;
END Set_Column;

PROCEDURE Set_Row (
   p_row       IN PLS_INTEGER,
   p_numFmtId  IN PLS_INTEGER  := null,
   p_fontId    IN PLS_INTEGER  := null,
   p_fillId    IN PLS_INTEGER  := null,
   p_borderId  IN PLS_INTEGER  := null,
   p_alignment IN tp_alignment := null,
   p_sheet     IN PLS_INTEGER  := null,
   p_height    IN NUMBER       := null )
IS
   t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
   t_cells tp_cells;
BEGIN
   workbook.sheets(t_sheet).row_fmts(p_row).numFmtId := p_numFmtId;
   workbook.sheets(t_sheet).row_fmts(p_row).fontId := p_fontId;
   workbook.sheets(t_sheet).row_fmts(p_row).fillId := p_fillId;
   workbook.sheets(t_sheet).row_fmts(p_row).borderId := p_borderId;
   workbook.sheets(t_sheet).row_fmts(p_row).alignment := p_alignment;
   workbook.sheets(t_sheet).row_fmts(p_row).height := trunc( p_height * 4 / 3 ) * 3 / 4;
   IF not workbook.sheets(t_sheet).rows.exists(p_row) THEN
      workbook.sheets(t_sheet).rows(p_row) := t_cells;
   END IF;
END Set_Row;

PROCEDURE Freeze_Rows (
   p_nr_rows IN PLS_INTEGER := 1,
   p_sheet   IN PLS_INTEGER := null )
IS
   t_sheet PLS_INTEGER := nvl(p_sheet, workbook.sheets.count());
BEGIN
   workbook.sheets(t_sheet).freeze_cols := null;
   workbook.sheets(t_sheet).freeze_rows := p_nr_rows;
END Freeze_Rows;
--
  PROCEDURE freeze_cols
    ( p_nr_cols PLS_INTEGER := 1
    , p_sheet PLS_INTEGER := null
    )
  is
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  begin
    workbook.sheets( t_sheet ).freeze_rows := null;
    workbook.sheets( t_sheet ).freeze_cols := p_nr_cols;
  end;
--
  PROCEDURE freeze_pane
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_sheet PLS_INTEGER := null
    )
  is
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  begin
    workbook.sheets( t_sheet ).freeze_rows := p_row;
    workbook.sheets( t_sheet ).freeze_cols := p_col;
  end;
--
  PROCEDURE set_autofilter
    ( p_column_start PLS_INTEGER := null
    , p_column_end PLS_INTEGER := null
    , p_row_start PLS_INTEGER := null
    , p_row_end PLS_INTEGER := null
    , p_sheet PLS_INTEGER := null
    )
  is
    t_ind PLS_INTEGER;
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  begin
    t_ind := 1;
    workbook.sheets( t_sheet ).autofilters( t_ind ).column_start := p_column_start;
    workbook.sheets( t_sheet ).autofilters( t_ind ).column_end := p_column_end;
    workbook.sheets( t_sheet ).autofilters( t_ind ).row_start := p_row_start;
    workbook.sheets( t_sheet ).autofilters( t_ind ).row_end := p_row_end;
    defined_name
      ( p_column_start
      , p_row_start
      , p_column_end
      , p_row_end
      , '_xlnm._FilterDatabase'
      , t_sheet
      , t_sheet - 1
      );
  end;
--
  PROCEDURE add1xml
    ( p_excel IN OUT NOCOPY BLOB
    , p_filename VARCHAR2
    , p_xml clob
    )
  is
    t_tmp blob;
    dest_offset integer := 1;
    src_offset integer := 1;
    lang_context integer;
    warning integer;
  begin
    lang_context := dbms_lob.DEFAULT_LANG_CTX;
    dbms_lob.createtemporary( t_tmp, true );
    dbms_lob.converttoblob
      ( t_tmp
      , p_xml
      , dbms_lob.lobmaxsize
      , dest_offset
      , src_offset
      ,  nls_charset_id( 'AL32UTF8'  )
      , lang_context
      , warning
      );
    add1file( p_excel, p_filename, t_tmp );
    dbms_lob.freetemporary( t_tmp );
  end;
--
FUNCTION Finish RETURN BLOB
IS
   t_excel   BLOB;
   t_yyy     BLOB;
   t_xxx     CLOB;
   t_tmp     VARCHAR2(32767 char);
   t_c       NUMBER;
   t_h       NUMBER;
   t_w       NUMBER;
   t_cw      NUMBER;
   s         PLS_INTEGER;
   t_row_ind PLS_INTEGER;
   t_col_min PLS_INTEGER;
   t_col_max PLS_INTEGER;
   t_col_ind PLS_INTEGER;
BEGIN
   dbms_lob.createtemporary(t_excel, true);
   t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
   s := workbook.sheets.first;
   WHILE s IS not null LOOP
      t_xxx := t_xxx || ( '
<Override PartName="/xl/worksheets/sheet' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' );
      s := workbook.sheets.next( s );
   END LOOP;
    t_xxx := t_xxx || '
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    s := workbook.sheets.first;
   WHILE s IS not null LOOP
      IF workbook.sheets( s ).comments.count() > 0 THEN
         t_xxx := t_xxx || ( '
<Override PartName="/xl/comments' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' );
      END IF;
      s := workbook.sheets.next( s );
   END LOOP;
   t_xxx := t_xxx || '
</Types>';
   add1xml( t_excel, '[Content_Types].xml', t_xxx );
   t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>' || sys_context( 'userenv', 'os_user' ) || '</dc:creator>
<dc:description>Build by version:' || VERSION_ || '</dc:description>
<cp:lastModifiedBy>' || sys_context( 'userenv', 'os_user' ) || '</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:modified>
</cp:coreProperties>';
    add1xml( t_excel, 'docProps/core.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
      t_xxx := t_xxx || ( '
<vt:lpstr>' || workbook.sheets(s).name || '</vt:lpstr>' );
      s := workbook.sheets.next(s);
   END LOOP;
   t_xxx := t_xxx || '</vt:vector>
</TitlesOfParts>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>14.0300</AppVersion>
</Properties>';
    add1xml( t_excel, 'docProps/app.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    add1xml( t_excel, '_rels/.rels', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
   IF workbook.numFmts.count() > 0 THEN
      t_xxx := t_xxx || ( '<numFmts count="' || workbook.numFmts.count() || '">' );
      FOR n IN 1 .. workbook.numFmts.count() LOOP
         t_xxx := t_xxx || ( '<numFmt numFmtId="' || workbook.numFmts( n ).numFmtId || '" formatCode="' || workbook.numFmts( n ).formatCode || '"/>' );
      END LOOP;
      t_xxx := t_xxx || '</numFmts>';
   END IF;
   t_xxx := t_xxx || ( '<fonts count="' || workbook.fonts.count() || '" x14ac:knownFonts="1">' );
   FOR f IN 0 .. workbook.fonts.count() - 1 LOOP
      t_xxx := t_xxx || ( '<font>' ||
         CASE WHEN workbook.fonts(f).bold THEN '<b/>' END ||
         CASE WHEN workbook.fonts(f).italic THEN '<i/>' END ||
         CASE WHEN workbook.fonts(f).underline THEN '<u/>' END ||
'<sz val="' || to_char( workbook.fonts( f ).fontsize, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )  || '"/>
<color ' || case when workbook.fonts( f ).rgb is not null
              then 'rgb="' || workbook.fonts( f ).rgb
              else 'theme="' || workbook.fonts( f ).theme
            end || '"/>
<name val="' || workbook.fonts( f ).name || '"/>
<family val="' || workbook.fonts( f ).family || '"/>
<scheme val="none"/>
</font>' );
   END LOOP;
    t_xxx := t_xxx || ( '</fonts>
<fills count="' || workbook.fills.count() || '">' );
    for f in 0 .. workbook.fills.count() - 1
    loop
      t_xxx := t_xxx || ( '<fill><patternFill patternType="' || workbook.fills( f ).patternType || '">' ||
         case when workbook.fills( f ).fgRGB is not null then '<fgColor rgb="' || workbook.fills( f ).fgRGB || '"/>' end ||
         case when workbook.fills( f ).bgRGB is not null then '<bgColor rgb="' || workbook.fills( f ).bgRGB || '"/>' end ||
         '</patternFill></fill>' );
    end loop;
    t_xxx := t_xxx || ( '</fills>
<borders count="' || workbook.borders.count() || '">' );
    for b in 0 .. workbook.borders.count() - 1
    loop
      t_xxx := t_xxx || ( '<border>' ||
         case when workbook.borders( b ).left   is null then '<left/>'   else '<left style="'   || workbook.borders( b ).left   || '"/>' end ||
         case when workbook.borders( b ).right  is null then '<right/>'  else '<right style="'  || workbook.borders( b ).right  || '"/>' end ||
         case when workbook.borders( b ).top    is null then '<top/>'    else '<top style="'    || workbook.borders( b ).top    || '"/>' end ||
         case when workbook.borders( b ).bottom is null then '<bottom/>' else '<bottom style="' || workbook.borders( b ).bottom || '"/>' end ||
         '</border>' );
    end loop;
    t_xxx := t_xxx || ( '</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="' || ( workbook.cellXfs.count() + 1 ) || '">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' );
    for x in 1 .. workbook.cellXfs.count()
    loop
      t_xxx := t_xxx || ( '<xf numFmtId="' || workbook.cellXfs( x ).numFmtId || '" fontId="' || workbook.cellXfs( x ).fontId || '" fillId="' || workbook.cellXfs( x ).fillId || '" borderId="' || workbook.cellXfs( x ).borderId || '">' );
      if (  workbook.cellXfs( x ).alignment.horizontal is not null
         or workbook.cellXfs( x ).alignment.vertical is not null
         or workbook.cellXfs( x ).alignment.wrapText
         )
      then
        t_xxx := t_xxx || ( '<alignment' ||
          case when workbook.cellXfs( x ).alignment.horizontal is not null then ' horizontal="' || workbook.cellXfs( x ).alignment.horizontal || '"' end ||
          case when workbook.cellXfs( x ).alignment.vertical is not null then ' vertical="' || workbook.cellXfs( x ).alignment.vertical || '"' end ||
          case when workbook.cellXfs( x ).alignment.wrapText then ' wrapText="true"' end || '/>' );
      end if;
      t_xxx := t_xxx || '</xf>';
    end loop;
    t_xxx := t_xxx || ( '</cellXfs>
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
    add1xml( t_excel, 'xl/styles.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr date1904="false" defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
    s := workbook.sheets.first;
    while s is not null
    loop
      t_xxx := t_xxx || ( '
<sheet name="' || workbook.sheets(s).name || '" sheetId="' || s || '" r:id="rId' || ( 9 + s ) || '"/>' );
      s := workbook.sheets.next(s);
    end loop;
    t_xxx := t_xxx || '</sheets>';
    if workbook.defined_names.count() > 0
    then
      t_xxx := t_xxx || '<definedNames>';
      for s in 1 .. workbook.defined_names.count()
      loop
        t_xxx := t_xxx || ( '
<definedName name="' || workbook.defined_names(s).name || '"' ||
            case when workbook.defined_names( s ).sheet is not null then ' localSheetId="' || to_char( workbook.defined_names( s ).sheet ) || '"' end ||
            '>' || workbook.defined_names( s ).ref || '</definedName>' );
      end loop;
      t_xxx := t_xxx || '</definedNames>';
    end if;
    t_xxx := t_xxx || '<calcPr calcId="144525"/></workbook>';
    add1xml( t_excel, 'xl/workbook.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    add1xml( t_excel, 'xl/theme/theme1.xml', t_xxx );
    s := workbook.sheets.first;
    while s is not null
    loop
      t_col_min := 16384;
      t_col_max := 1;
      t_row_ind := workbook.sheets( s ).rows.first();
      while t_row_ind is not null
      loop
        t_col_min := least( t_col_min, workbook.sheets( s ).rows( t_row_ind ).first() );
        t_col_max := greatest( t_col_max, workbook.sheets( s ).rows( t_row_ind ).last() );
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      end loop;
      addtxt2utf8blob_init( t_yyy );
      addtxt2utf8blob( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' ||
case when workbook.sheets( s ).tabcolor is not null then '<sheetPr><tabColor rgb="' || workbook.sheets( s ).tabcolor || '"/></sheetPr>' end ||
'<dimension ref="' || alfan_col( t_col_min ) || workbook.sheets( s ).rows.first() || ':' || alfan_col( t_col_max ) || workbook.sheets( s ).rows.last() || '"/>
<sheetViews>
<sheetView' || case when s = 1 then ' tabSelected="1"' end || ' workbookViewId="0">'
                     , t_yyy
                     );
      if workbook.sheets( s ).freeze_rows > 0 and workbook.sheets( s ).freeze_cols > 0
      then
        addtxt2utf8blob( '<pane xSplit="' || workbook.sheets( s ).freeze_cols || '" '
                          || 'ySplit="' || workbook.sheets( s ).freeze_rows || '" '
                          || 'topLeftCell="' || alfan_col( workbook.sheets( s ).freeze_cols + 1 ) || ( workbook.sheets( s ).freeze_rows + 1 ) || '" '
                          || 'activePane="bottomLeft" state="frozen"/>'
                       , t_yyy
                       );
      else
        if workbook.sheets( s ).freeze_rows > 0
        then
          addtxt2utf8blob( '<pane ySplit="' || workbook.sheets( s ).freeze_rows || '" topLeftCell="A' || ( workbook.sheets( s ).freeze_rows + 1 ) || '" activePane="bottomLeft" state="frozen"/>'
                         , t_yyy
                         );
        end if;
        if workbook.sheets( s ).freeze_cols > 0
        then
          addtxt2utf8blob( '<pane xSplit="' || workbook.sheets( s ).freeze_cols || '" topLeftCell="' || alfan_col( workbook.sheets( s ).freeze_cols + 1 ) || '1" activePane="bottomLeft" state="frozen"/>'
                         , t_yyy
                         );
        end if;
      end if;
      addtxt2utf8blob( '</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
                     , t_yyy
                     );
      if workbook.sheets( s ).widths.count() > 0
      then
        addtxt2utf8blob( '<cols>', t_yyy );
        t_col_ind := workbook.sheets( s ).widths.first();
        while t_col_ind is not null
        loop
          addtxt2utf8blob( '<col min="' || t_col_ind || '" max="' || t_col_ind || '" width="' || to_char( workbook.sheets( s ).widths( t_col_ind ), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" customWidth="1"/>', t_yyy );
          t_col_ind := workbook.sheets( s ).widths.next( t_col_ind );
        end loop;
        addtxt2utf8blob( '</cols>', t_yyy );
      end if;
      addtxt2utf8blob( '<sheetData>', t_yyy );
      t_row_ind := workbook.sheets( s ).rows.first();
      while t_row_ind is not null
      loop
        if workbook.sheets( s ).row_fmts.exists( t_row_ind ) and workbook.sheets( s ).row_fmts( t_row_ind ).height is not null
        then
          addtxt2utf8blob( '<row r="' || t_row_ind || '" spans="' || t_col_min || ':' || t_col_max || '" customHeight="1" ht="'
                         || to_char( workbook.sheets( s ).row_fmts( t_row_ind ).height, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" >', t_yyy );
        else
          addtxt2utf8blob( '<row r="' || t_row_ind || '" spans="' || t_col_min || ':' || t_col_max || '">', t_yyy );
        end if;
        t_col_ind := workbook.sheets( s ).rows( t_row_ind ).first();
        while t_col_ind is not null
        loop
          if workbook.sheets( s ).rows( t_row_ind )( t_col_ind ).formula_idx is null
          then
            t_tmp := null;
          else
            t_tmp := '<f>' || workbook.formulas( workbook.sheets( s ).rows( t_row_ind )( t_col_ind ).formula_idx ) || '</f>';
          end if;
          addtxt2utf8blob( '<c r="' || alfan_col( t_col_ind ) || t_row_ind || '"'
                 || ' ' || workbook.sheets( s ).rows( t_row_ind )( t_col_ind ).style
                 || '>' || t_tmp || '<v>'
                 || to_char( workbook.sheets( s ).rows( t_row_ind )( t_col_ind ).value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )
                 || '</v></c>', t_yyy );
          t_col_ind := workbook.sheets( s ).rows( t_row_ind ).next( t_col_ind );
        end loop;
        addtxt2utf8blob( '</row>', t_yyy );
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      end loop;
      addtxt2utf8blob( '</sheetData>', t_yyy );
      for a in 1 ..  workbook.sheets( s ).autofilters.count()
      loop
        addtxt2utf8blob( '<autoFilter ref="' ||
            alfan_col( nvl( workbook.sheets( s ).autofilters( a ).column_start, t_col_min ) ) ||
            nvl( workbook.sheets( s ).autofilters( a ).row_start, workbook.sheets( s ).rows.first() ) || ':' ||
            alfan_col( coalesce( workbook.sheets( s ).autofilters( a ).column_end, workbook.sheets( s ).autofilters( a ).column_start, t_col_max ) ) ||
            nvl( workbook.sheets( s ).autofilters( a ).row_end, workbook.sheets( s ).rows.last() ) || '"/>', t_yyy );
      end loop;
      if workbook.sheets( s ).mergecells.count() > 0
      then
        addtxt2utf8blob( '<mergeCells count="' || to_char( workbook.sheets( s ).mergecells.count() ) || '">', t_yyy );
        for m in 1 ..  workbook.sheets( s ).mergecells.count()
        loop
          addtxt2utf8blob( '<mergeCell ref="' || workbook.sheets( s ).mergecells( m ) || '"/>', t_yyy );
        end loop;
        addtxt2utf8blob( '</mergeCells>', t_yyy );
      end if;
--
      if workbook.sheets( s ).validations.count() > 0
      then
        addtxt2utf8blob( '<dataValidations count="' || to_char( workbook.sheets( s ).validations.count() ) || '">', t_yyy );
        for m in 1 ..  workbook.sheets( s ).validations.count()
        loop
          addtxt2utf8blob( '<dataValidation' ||
              ' type="' || workbook.sheets( s ).validations( m ).type || '"' ||
              ' errorStyle="' || workbook.sheets( s ).validations( m ).errorstyle || '"' ||
              ' allowBlank="' || case when nvl( workbook.sheets( s ).validations( m ).allowBlank, true ) then '1' else '0' end || '"' ||
              ' sqref="' || workbook.sheets( s ).validations( m ).sqref || '"', t_yyy );
          if workbook.sheets( s ).validations( m ).prompt is not null
          then
            addtxt2utf8blob( ' showInputMessage="1" prompt="' || workbook.sheets( s ).validations( m ).prompt || '"', t_yyy );
            if workbook.sheets( s ).validations( m ).title is not null
            then
              addtxt2utf8blob( ' promptTitle="' || workbook.sheets( s ).validations( m ).title || '"', t_yyy );
            end if;
          end if;
          if workbook.sheets( s ).validations( m ).showerrormessage
          then
            addtxt2utf8blob( ' showErrorMessage="1"', t_yyy );
            if workbook.sheets( s ).validations( m ).error_title is not null
            then
              addtxt2utf8blob( ' errorTitle="' || workbook.sheets( s ).validations( m ).error_title || '"', t_yyy );
            end if;
            if workbook.sheets( s ).validations( m ).error_txt is not null
            then
              addtxt2utf8blob( ' error="' || workbook.sheets( s ).validations( m ).error_txt || '"', t_yyy );
            end if;
          end if;
          addtxt2utf8blob( '>', t_yyy );
          if workbook.sheets( s ).validations( m ).formula1 is not null
          then
            addtxt2utf8blob( '<formula1>' || workbook.sheets( s ).validations( m ).formula1 || '</formula1>', t_yyy );
          end if;
          if workbook.sheets( s ).validations( m ).formula2 is not null
          then
            addtxt2utf8blob( '<formula2>' || workbook.sheets( s ).validations( m ).formula2 || '</formula2>', t_yyy );
          end if;
          addtxt2utf8blob( '</dataValidation>', t_yyy );
        end loop;
        addtxt2utf8blob( '</dataValidations>', t_yyy );
      end if;
--
      if workbook.sheets( s ).hyperlinks.count() > 0
      then
        addtxt2utf8blob( '<hyperlinks>', t_yyy );
        for h in 1 ..  workbook.sheets( s ).hyperlinks.count()
        loop
          addtxt2utf8blob( '<hyperlink ref="' || workbook.sheets( s ).hyperlinks( h ).cell || '" r:id="rId' || h || '"/>', t_yyy );
        end loop;
        addtxt2utf8blob( '</hyperlinks>', t_yyy );
      end if;
      addtxt2utf8blob( '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', t_yyy );
      if workbook.sheets( s ).comments.count() > 0
      then
        addtxt2utf8blob( '<legacyDrawing r:id="rId' || ( workbook.sheets( s ).hyperlinks.count() + 1 ) || '"/>', t_yyy );
      end if;
--
      addtxt2utf8blob( '</worksheet>', t_yyy );
      addtxt2utf8blob_finish( t_yyy );
      add1file( t_excel, 'xl/worksheets/sheet' || s || '.xml', t_yyy );
      if workbook.sheets( s ).hyperlinks.count() > 0 or workbook.sheets( s ).comments.count() > 0
      then
        t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        if workbook.sheets( s ).comments.count() > 0
        then
          t_xxx := t_xxx || ( '<Relationship Id="rId' || ( workbook.sheets( s ).hyperlinks.count() + 2 ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' || s || '.xml"/>' );
          t_xxx := t_xxx || ( '<Relationship Id="rId' || ( workbook.sheets( s ).hyperlinks.count() + 1 ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' || s || '.vml"/>' );
        end if;
        for h in 1 ..  workbook.sheets( s ).hyperlinks.count()
        loop
          t_xxx := t_xxx || ( '<Relationship Id="rId' || h || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' || workbook.sheets( s ).hyperlinks( h ).url || '" TargetMode="External"/>' );
        end loop;
        t_xxx := t_xxx || '</Relationships>';
        add1xml( t_excel, 'xl/worksheets/_rels/sheet' || s || '.xml.rels', t_xxx );
      end if;
--
      if workbook.sheets( s ).comments.count() > 0
      then
        declare
          cnt PLS_INTEGER;
          author_ind tp_author;
--          t_col_ind := workbook.sheets( s ).widths.next( t_col_ind );
        begin
          authors.delete();
          for c in 1 .. workbook.sheets( s ).comments.count()
          loop
            authors( workbook.sheets( s ).comments( c ).author ) := 0;
          end loop;
          t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>';
          cnt := 0;
          author_ind := authors.first();
          while author_ind is not null or authors.next( author_ind ) is not null
          loop
            authors( author_ind ) := cnt;
            t_xxx := t_xxx || ( '<author>' || author_ind || '</author>' );
            cnt := cnt + 1;
            author_ind := authors.next( author_ind );
          end loop;
        end;
        t_xxx := t_xxx || '</authors><commentList>';
        for c in 1 .. workbook.sheets( s ).comments.count()
        loop
          t_xxx := t_xxx || ( '<comment ref="' || alfan_col( workbook.sheets( s ).comments( c ).column ) ||
             to_char( workbook.sheets( s ).comments( c ).row || '" authorId="' || authors( workbook.sheets( s ).comments( c ).author ) ) || '">
<text>' );
          if workbook.sheets( s ).comments( c ).author is not null
          then
            t_xxx := t_xxx || ( '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
               workbook.sheets( s ).comments( c ).author || ':</t></r>' );
          end if;
          t_xxx := t_xxx || ( '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
             case when workbook.sheets( s ).comments( c ).author is not null then '
' end || workbook.sheets( s ).comments( c ).text || '</t></r></text></comment>' );
        end loop;
        t_xxx := t_xxx || '</commentList></comments>';
        add1xml( t_excel, 'xl/comments' || s || '.xml', t_xxx );
        t_xxx := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
        for c in 1 .. workbook.sheets( s ).comments.count()
        loop
          t_xxx := t_xxx || ( '<v:shape id="_x0000_s' || to_char( c ) || '" type="#_x0000_t202"
style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' || to_char( c ) || ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' );
          t_w := workbook.sheets( s ).comments( c ).width;
          t_c := 1;
          loop
            if workbook.sheets( s ).widths.exists( workbook.sheets( s ).comments( c ).column + t_c )
            then
              t_cw := 256 * workbook.sheets( s ).widths( workbook.sheets( s ).comments( c ).column + t_c );
              t_cw := trunc( ( t_cw + 18 ) / 256 * 7); -- assume default 11 point Calibri
            else
              t_cw := 64;
            end if;
            exit when t_w < t_cw;
            t_c := t_c + 1;
            t_w := t_w - t_cw;
          end loop;
          t_h := workbook.sheets( s ).comments( c ).height;
          t_xxx := t_xxx || ( '<x:Anchor>' || workbook.sheets( s ).comments( c ).column || ',15,' ||
                     workbook.sheets( s ).comments( c ).row || ',30,' ||
                     ( workbook.sheets( s ).comments( c ).column + t_c - 1 ) || ',' || round( t_w ) || ',' ||
                     ( workbook.sheets( s ).comments( c ).row + 1 + trunc( t_h / 20 ) ) || ',' || mod( t_h, 20 ) || '</x:Anchor>' );
          t_xxx := t_xxx || ( '<x:AutoFill>False</x:AutoFill><x:Row>' ||
            ( workbook.sheets( s ).comments( c ).row - 1 ) || '</x:Row><x:Column>' ||
            ( workbook.sheets( s ).comments( c ).column - 1 ) || '</x:Column></x:ClientData></v:shape>' );
        end loop;
        t_xxx := t_xxx || '</xml>';
        add1xml( t_excel, 'xl/drawings/vmlDrawing' || s || '.vml', t_xxx );
      end if;
--
      s := workbook.sheets.next( s );
    end loop;
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
    s := workbook.sheets.first;
    while s is not null
    loop
      t_xxx := t_xxx || ( '
<Relationship Id="rId' || ( 9 + s ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || s || '.xml"/>' );
      s := workbook.sheets.next( s );
    end loop;
    t_xxx := t_xxx || '</Relationships>';
    add1xml( t_excel, 'xl/_rels/workbook.xml.rels', t_xxx );
    addtxt2utf8blob_init( t_yyy );
    addtxt2utf8blob( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || workbook.str_cnt || '" uniqueCount="' || workbook.strings.count() || '">'
                  , t_yyy
                  );
    for i in 0 .. workbook.str_ind.count() - 1
    loop
      addtxt2utf8blob( '<si><t xml:space="preserve">' || dbms_xmlgen.convert( substr( workbook.str_ind( i ), 1, 32000 ) ) || '</t></si>', t_yyy );
    end loop;
    addtxt2utf8blob( '</sst>', t_yyy );
    addtxt2utf8blob_finish( t_yyy );
    add1file( t_excel, 'xl/sharedStrings.xml', t_yyy );
    finish_zip( t_excel );
    clear_workbook;
    return t_excel;
END Finish;

PROCEDURE Save (
   p_directory IN VARCHAR2,
   p_filename  IN VARCHAR2 )
IS BEGIN
   blob2file (finish, p_directory, p_filename);
END Save;

PROCEDURE query2sheet (
   cur_             IN OUT INTEGER,
   p_column_headers IN BOOLEAN     := true,
   p_directory      IN VARCHAR2    := null,
   p_filename       IN VARCHAR2    := null,
   p_sheet          IN PLS_INTEGER := null,
   p_UseXf          IN BOOLEAN     := false )
IS
   t_sheet     PLS_INTEGER;
   t_col_cnt   INTEGER;
   t_desc_tab  dbms_sql.desc_tab2;
   d_tab       dbms_sql.date_table;
   n_tab       dbms_sql.number_table;
   v_tab       dbms_sql.VARCHAR2_table;
   t_bulk_size PLS_INTEGER := 200;
   rows_       INTEGER;
   t_cur_row   PLS_INTEGER;
   t_useXf     BOOLEAN := g_useXf;
   TYPE tp_XfIds IS TABLE OF VARCHAR2(50) INDEX BY PLS_INTEGER;
   t_XfIds     tp_XfIds;
BEGIN
   IF p_sheet IS null THEN
      New_Sheet;
   END IF;
   t_sheet := coalesce (p_sheet, workbook.sheets.count());
   setUseXf(true);
   dbms_sql.describe_columns2(cur_, t_col_cnt, t_desc_tab);
   FOR col_ IN 1 .. t_col_cnt LOOP
      IF p_column_headers THEN
         Cell (col_, 1, t_desc_tab(col_).col_name, p_sheet => t_sheet );
      END IF;
      CASE
         WHEN t_desc_tab(col_).col_type IN (2, 100, 101) THEN
            dbms_sql.define_array (cur_, col_, n_tab, t_bulk_size, 1);
         WHEN t_desc_tab(col_).col_type IN (12, 178, 179, 180, 181, 231) THEN
            dbms_sql.define_array (cur_, col_, d_tab, t_bulk_size, 1);
            t_XfIds(col_) := get_XfId( t_sheet, col_, null, get_numFmt('dd/mm/yyyy'));
         WHEN t_desc_tab(col_).col_type IN (1, 8, 9, 96, 112) THEN
            dbms_sql.define_array (cur_, col_, v_tab, t_bulk_size, 1);
         ELSE
            null;
      END CASE;
   END LOOP;
   setUseXf (p_UseXf);
   t_cur_row := CASE WHEN p_column_headers THEN 2 ELSE 1 END;

   LOOP
      rows_ := dbms_sql.fetch_rows(cur_);
      IF rows_ > 0 THEN
         FOR col_ IN 1 .. t_col_cnt LOOP
            CASE
               WHEN t_desc_tab(col_).col_type IN (2, 100, 101) THEN
                  dbms_sql.column_value (cur_, col_, n_tab);
                  FOR i_ IN 0 .. rows_ - 1 LOOP
                     IF n_tab(i_+n_tab.first() ) IS NOT null THEN
                        Cell(col_, t_cur_row+i_, n_tab(i_+n_tab.first()), p_sheet => t_sheet);
                     END IF;
                  END LOOP;
                  n_tab.delete;
               WHEN t_desc_tab(col_).col_type IN (12, 178, 179, 180, 181, 231) THEN
                  dbms_sql.column_value(cur_, col_, d_tab);
                  FOR i_ IN 0 .. rows_ - 1 LOOP
                     IF d_tab(i_+d_tab.first()) IS NOT null THEN
                        IF g_useXf THEN
                           Cell (col_, t_cur_row+i_, d_tab(i_+d_tab.first()), p_sheet => t_sheet);
                        ELSE
                           query_date_cell(col_, t_cur_row+i_, d_tab(i_+d_tab.first()), t_sheet, t_XfIds(col_));
                        END IF;
                     END IF;
                  END LOOP;
                  d_tab.delete;
               WHEN t_desc_tab(col_).col_type IN (1, 8, 9, 96, 112) THEN
                  dbms_sql.column_value (cur_, col_, v_tab);
                  FOR i_ IN 0 .. rows_ - 1 LOOP
                     IF v_tab(i_+v_tab.first()) IS NOT null THEN
                        Cell (col_, t_cur_row+i_, v_tab(i_+v_tab.first()), p_sheet => t_sheet);
                     END IF;
                  END LOOP;
                  v_tab.delete;
               ELSE
                  null;
            END CASE;
         END LOOP;
      END IF;
      EXIT WHEN rows_ != t_bulk_size;
      t_cur_row := t_cur_row + rows_;
   END LOOP; -- loop for each column in the result set
   dbms_sql.close_cursor (cur_);
   IF p_directory IS NOT null AND p_filename IS NOT null THEN
      Save (p_directory, p_filename);
   END IF;
   setUseXf (t_useXf);
EXCEPTION
   WHEN others THEN
      IF dbms_sql.is_open (cur_) THEN
         dbms_sql.close_cursor (cur_);
      END IF;
      setUseXf(t_useXf);
END query2sheet;

PROCEDURE query2sheet (
   sql_            IN VARCHAR2,
   column_headers_ IN BOOLEAN     := true,
   directory_      IN VARCHAR2    := null,
   filename_       IN VARCHAR2    := null,
   sheet_          IN PLS_INTEGER := null,
   UseXf_          IN BOOLEAN     := false )
IS
   cur_ INTEGER := Dbms_Sql.Open_Cursor;
   res_ INTEGER;
BEGIN
   Dbms_Sql.Parse (cur_, sql_, dbms_sql.native);
   res_ := Dbms_Sql.Execute(cur_);
   Dbms_Output.Put_Line ('rows of data returned by cursor: ' || res_);
   query2sheet (cur_, column_headers_, directory_, filename_, sheet_, UseXf_);
END query2sheet;

PROCEDURE query2sheet (
   p_rc             IN OUT SYS_REFCURSOR,
   p_column_headers IN BOOLEAN     := true,
   p_directory      IN VARCHAR2    := null,
   p_filename       IN VARCHAR2    := null,
   p_sheet          IN PLS_INTEGER := null,
   p_UseXf          IN BOOLEAN := false )
IS
   cur_ INTEGER := dbms_sql.to_cursor_number (p_rc);
BEGIN
   query2sheet (
      cur_             => cur_,
      p_column_headers => p_column_headers,
      p_directory      => p_directory,
      p_filename       => p_filename,
      p_sheet          => p_sheet,
      p_UseXf          => p_UseXf
   );
END query2sheet;


PROCEDURE setUseXf (
   p_val BOOLEAN := true )
IS BEGIN
   g_useXf := p_val;
END setUseXf;


BEGIN
   Clear_Workbook;
   New_Sheet ('Sheet 1');
END AS_XLSX;
/
