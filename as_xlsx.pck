CREATE OR REPLACE PACKAGE as_xlsx IS
/*****************************************************************************
 *****************************************************************************
 **
 ** Author: Anton Scheffer
 ** Website: http://technology.amis.nl/blog
 ** See also: http://technology.amis.nl/blog/?p=10995
 **   # License
 **   Copyright (C) 2011, 2020 by Anton Scheffer
 **   See associated LICENSE.md file
 **
 ** Modifications added by Osian ap Garth since 2017, version-controlled since
 ** 2021 in Git Hub:
 **    >> https://github.com/cartbeforehorse/as_xlsx
 ** Documentation updated in README.md
 **
 *****************************************************************************
 *****************************************************************************/


--------------------------------------------------
-- Public Types
--

-----
-- Excel cell structure, cell and range locators
--
TYPE tp_pivot_cols  IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;
TYPE tp_col_agg_fns IS TABLE OF VARCHAR2(20) INDEX BY PLS_INTEGER; -- [sum,avg,count...]
TYPE tp_pivot_axes IS RECORD (
   vrollups    tp_pivot_cols,
   hrollups    tp_pivot_cols,
   filter_cols tp_pivot_cols,
   col_agg_fns tp_col_agg_fns
);
TYPE tp_cell_loc IS RECORD (
   c     PLS_INTEGER,         -- 2
   r     PLS_INTEGER,         -- 3
   fixc  BOOLEAN  := false,
   fixr  BOOLEAN  := false ); -- true ==> B$3

TYPE tp_column_names  IS TABLE OF VARCHAR2(2000) INDEX BY PLS_INTEGER;
TYPE tp_cell_range IS RECORD (
   sheet_id     PLS_INTEGER,   -- sheet.name => My Perfect Sheet; nullable, a range doens't necessarily need a sheet
   tl           tp_cell_loc,   -- (2, 3, false, true)
   br           tp_cell_loc,   -- (6, 6, false, false) Alfan_Range() => 'My Perfect Sheet'!B$3:F6
   defined_name VARCHAR2(100), -- 'MyDatacells'
   local_sheet  BOOLEAN,       -- sets the defined name to be accessible only on `sheet_id`
   col_names    tp_column_names ); -- makes our lives easier in building pivots

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
TYPE border_list IS TABLE OF INTEGER INDEX BY VARCHAR2(50);
TYPE numFmt_list IS TABLE OF INTEGER INDEX BY VARCHAR2(50);
TYPE align_list  IS TABLE OF tp_alignment INDEX BY VARCHAR2(50);
TYPE numFmt_cols IS TABLE OF INTEGER INDEX BY PLS_INTEGER;

fonts_  fonts_list;
fills_  fills_list;
bdrs_   border_list;
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
   format_mask_ IN VARCHAR2 := null ) RETURN PLS_INTEGER;

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

PROCEDURE Add_Fill (
   fill_id_     IN VARCHAR2,
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,
   bgRGB_       IN VARCHAR2 := null );

PROCEDURE Add_NumFmt (
   fmt_id_ IN VARCHAR2,
   format_ IN VARCHAR2 );

PROCEDURE Print_Range (
   range_ IN tp_cell_range );

---------------------------------------
-- Alfan_Cell(), Alfan_Range()
--  Transforms a numeric cell or range reference into an Excel reference.  For
--  example [1, 2] becomes "A2"; [1, 2, 3, 8] becomes "A2:C8".  This is useful
--  when external code is trying to generate formulas.
--
FUNCTION Alfan_Cell (
   col_  IN PLS_INTEGER,
   row_  IN PLS_INTEGER,
   fix1_ IN BOOLEAN := false,
   fix2_ IN BOOLEAN := false ) RETURN VARCHAR2;

FUNCTION Alfan_Range (
   col_tl_  IN PLS_INTEGER,
   row_tl_  IN PLS_INTEGER,
   col_br_  IN PLS_INTEGER,
   row_br_  IN PLS_INTEGER,
   fix_tlc_ IN BOOLEAN := false,
   fix_tlr_ IN BOOLEAN := false,
   fix_brc_ IN BOOLEAN := false,
   fix_brr_ IN BOOLEAN := false ) RETURN VARCHAR2;

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

PROCEDURE Add_Border_To_Range (
   cell_left_ IN PLS_INTEGER,
   cell_top_  IN PLS_INTEGER,
   width_     IN PLS_INTEGER,
   height_    IN PLS_INTEGER,
   style_     IN VARCHAR2    := 'medium',
   sheet_     IN PLS_INTEGER := null );

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
PROCEDURE Cell (
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_num_ IN NUMBER      := null,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null );
PROCEDURE CellN ( -- num version overload
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_num_ IN NUMBER      := null,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null );

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
PROCEDURE Cell (
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_str_ IN VARCHAR2    := '',
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null );
PROCEDURE CellS ( -- string version overload
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_str_ IN VARCHAR2    := '',
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null );

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
PROCEDURE Cell (
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_dt_  IN DATE        := null,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null );
PROCEDURE CellD ( -- date version overload
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_dt_  IN DATE,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null );

PROCEDURE CellB ( -- empty
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null );

PROCEDURE Condition_Color_Col (
   col_   IN PLS_INTEGER,
   sheet_ IN PLS_INTEGER := null );

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
   sheet_         IN PLS_INTEGER  := null );

PROCEDURE Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN NUMBER      := null,
   numFmtId_      IN VARCHAR2    := null,
   fontId_        IN VARCHAR2    := null,
   fillId_        IN VARCHAR2    := null,
   borderId_      IN VARCHAR2    := null,
   alignment_     IN VARCHAR2    := null,
   sheet_         IN PLS_INTEGER := null );

PROCEDURE Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN VARCHAR2    := null,
   numFmtId_      IN VARCHAR2    := null,
   fontId_        IN VARCHAR2    := null,
   fillId_        IN VARCHAR2    := null,
   borderId_      IN VARCHAR2    := null,
   alignment_     IN VARCHAR2    := null,
   sheet_         IN PLS_INTEGER := null );

PROCEDURE Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN DATE        := null,
   numFmtId_      IN VARCHAR2    := null,
   fontId_        IN VARCHAR2    := null,
   fillId_        IN VARCHAR2    := null,
   borderId_      IN VARCHAR2    := null,
   alignment_     IN VARCHAR2    := null,
   sheet_         IN PLS_INTEGER := null );

PROCEDURE Mergecells (
   tl_col_ IN PLS_INTEGER, -- top left
   tl_row_ IN PLS_INTEGER,
   br_col_ IN PLS_INTEGER, -- bottom right
   br_row_ IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null );

PROCEDURE List_Validation (
   sqref_col_   IN PLS_INTEGER,
   sqref_row_   IN PLS_INTEGER,
   tl_col_      IN PLS_INTEGER, -- top left
   tl_row_      IN PLS_INTEGER,
   br_col_      IN PLS_INTEGER, -- bottom right
   br_row_      IN PLS_INTEGER,
   style_       IN VARCHAR2    := 'stop', -- stop, warning, information
   title_       IN VARCHAR2    := null,
   prompt_      IN VARCHAR     := null,
   show_error_  IN BOOLEAN     := false,
   error_title_ IN VARCHAR2    := null,
   error_txt_   IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null );

PROCEDURE List_Validation (
   sqref_col_    IN PLS_INTEGER,
   sqref_row_    IN PLS_INTEGER,
   defined_name_ IN VARCHAR2,
   style_        IN VARCHAR2    := 'stop', -- stop, warning, information
   title_        IN VARCHAR2    := null,
   prompt_       IN VARCHAR     := null,
   show_error_   IN BOOLEAN     := false,
   error_title_  IN VARCHAR2    := null,
   error_txt_    IN VARCHAR2    := null,
   sheet_        IN PLS_INTEGER := null );

PROCEDURE Add_Image (
   col_         IN PLS_INTEGER,
   row_         IN PLS_INTEGER,
   img_blob_    IN BLOB,
   name_        IN VARCHAR2    := '',
   title_       IN VARCHAR2    := '',
   description_ IN VARCHAR2    := '',
   scale_       IN NUMBER      := null,
   sheet_       IN PLS_INTEGER := null,
   width_       IN PLS_INTEGER := null,
   height_      IN PLS_INTEGER := null );

PROCEDURE Load_Image (
   col_         IN PLS_INTEGER,
   row_         IN PLS_INTEGER,
   dir_         IN VARCHAR2,
   filename_    IN VARCHAR2,
   name_        IN VARCHAR2    := '',
   title_       IN VARCHAR2    := '',
   description_ IN VARCHAR2    := '',
   scale_       IN NUMBER      := null,
   sheet_       IN PLS_INTEGER := null,
   width_       IN PLS_INTEGER := null,
   height_      IN PLS_INTEGER := null );

PROCEDURE Defined_Name (
   name_       VARCHAR2,
   tl_col_     PLS_INTEGER, -- top left
   tl_row_     PLS_INTEGER,
   br_col_     PLS_INTEGER, -- bottom right
   br_row_     PLS_INTEGER,
   fix_tlc_    BOOLEAN     := true,
   fix_tlr_    BOOLEAN     := true,
   fix_brc_    BOOLEAN     := true,
   fix_brr_    BOOLEAN     := true,
   sheet_      PLS_INTEGER := null,
   localsheet_ BOOLEAN     := false );

PROCEDURE Defined_Name (
   range_ IN tp_cell_range );

FUNCTION Range_From_Defined_Name (
   defined_name_ IN VARCHAR2 ) RETURN tp_cell_range;

FUNCTION Add_Pivot_Cache (
   src_data_range_ IN OUT NOCOPY tp_cell_range,
   pivot_axes_     IN tp_pivot_axes ) RETURN PLS_INTEGER;

PROCEDURE Add_Pivot_Table (
   cache_id_       IN OUT NOCOPY PLS_INTEGER,
   src_data_range_ IN OUT NOCOPY tp_cell_range,
   pivot_axes_     IN tp_pivot_axes,
   location_tl_    IN tp_cell_loc,
   pivot_name_     IN VARCHAR2    := null,
   add_to_sheet_   IN PLS_INTEGER := null,
   new_sheet_name_ IN VARCHAR2    := null );

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
   name_   IN VARCHAR2,
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


END as_xlsx;
/
CREATE OR REPLACE PACKAGE BODY as_xlsx IS

VERSION_ CONSTANT VARCHAR2(20) := 'as_xlsx20';

LOCAL_FILE_HEADER_        CONSTANT RAW(4) := hextoraw('504B0304'); -- Local file header signature
END_OF_CENTRAL_DIRECTORY_ CONSTANT RAW(4) := hextoraw('504B0506'); -- End of central directory signature

-- Excel's first day is 1900-01-01, represented by the number 1 (not 0), which
-- means we need to subtract 1 day from 01-01-1900.  What's more is that Excel
-- thinks that 1900 is a leap year, and so from 1900-03-01 we need to take off
-- another day, leaving us with a base-date of 1899-12-30:
EXCEL_BASE_DATE_          CONSTANT DATE         := to_date('18991230','YYYYMMDD');
CELL_DT_STRING_           CONSTANT VARCHAR2(10) := 'string';
CELL_DT_NUMBER_           CONSTANT VARCHAR2(10) := 'number';
CELL_DT_DATE_             CONSTANT VARCHAR2(10) := 'date';
--CELL_DT_BOOL_             CONSTANT VARCHAR2(10) := 'bool';
CELL_DT_HYPERLINK_        CONSTANT VARCHAR2(10) := 'hyperlink';


---------------------------------------
---------------------------------------
--
-- Type Definitions
--
--

-----
-- formatting types
--
TYPE tp_XF_fmt IS RECORD (
   numFmtId  PLS_INTEGER,
   fontId    PLS_INTEGER,
   fillId    PLS_INTEGER,
   borderId  PLS_INTEGER,
   alignment tp_alignment,
   height    NUMBER,
   md5       RAW(128)
);
TYPE tp_col_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
TYPE tp_row_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
TYPE tp_widths IS TABLE OF NUMBER INDEX BY PLS_INTEGER;

-----
-- Excel cell structure
--

-- Cell properties
TYPE tp_cell_value IS RECORD (
   str_val  VARCHAR2(32000),
   num_val  NUMBER,
   dt_val   DATE,   -- dates are stored as numbers in Excel, but this is convenient
   bl_val   BOOLEAN -- not yet implemented as a cell type
);
TYPE tp_cell IS RECORD (
   datatype    VARCHAR2(30), -- string|number|date|bool|hyperlink
   ora_value   tp_cell_value,
   value       NUMBER,
   style       PLS_INTEGER,
   formula_idx PLS_INTEGER
);
TYPE tp_cells IS TABLE OF tp_cell INDEX BY PLS_INTEGER;
TYPE tp_rows IS TABLE OF tp_cells INDEX BY PLS_INTEGER;

TYPE tp_autofilter IS RECORD (
   column_start PLS_INTEGER,
   column_end   PLS_INTEGER,
   row_start    PLS_INTEGER,
   row_end      PLS_INTEGER
);
TYPE tp_autofilters IS TABLE OF tp_autofilter INDEX BY PLS_INTEGER;

TYPE tp_hyperlink IS RECORD (
   cell VARCHAR2(10),
   url  VARCHAR2(1000)
);
TYPE tp_hyperlinks IS TABLE OF tp_hyperlink INDEX BY PLS_INTEGER;

-----
-- comment types
SUBTYPE tp_author IS VARCHAR2(32767 char);
TYPE tp_authors IS TABLE OF PLS_INTEGER INDEX BY tp_author;

TYPE tp_comment IS RECORD (
   text   VARCHAR2(32767 char),
   author tp_author,
   row    PLS_INTEGER,
   column PLS_INTEGER,
   width  PLS_INTEGER,
   height PLS_INTEGER
);
TYPE tp_comments   IS TABLE OF tp_comment INDEX BY PLS_INTEGER;

TYPE tp_mergecells IS TABLE OF VARCHAR2(21) INDEX BY PLS_INTEGER;

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

-----
-- pivot types

-----
-- tp_unique_data  =>
-- tp_data_ix_ord =>
--   Sometimes we want to store data indexed by a unique string value.  But as
--   other times we need to keep that data ordered, which doesn't suit the way
--   that plsql-table-types work.  We therefore hold the data in two different
--   arrays, defined by these types.
--
TYPE tp_unique_data IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(32000);
TYPE tp_data_ix_ord IS TABLE OF VARCHAR2(32000) INDEX BY PLS_INTEGER;
TYPE tp_col_filters IS TABLE OF VARCHAR2(32000) INDEX BY PLS_INTEGER;

TYPE tp_cache_field IS RECORD (
   field_name   VARCHAR2(2000),
   rollup_fn    VARCHAR2(20),
   format_id    PLS_INTEGER,
   shared_items tp_unique_data,
   si_order     tp_data_ix_ord,
   min_value    NUMBER,
   max_value    NUMBER
);
TYPE tp_cache_fields IS TABLE OF tp_cache_field INDEX BY VARCHAR2(32000);
TYPE tp_pivot_cache IS RECORD (
   cache_id       PLS_INTEGER,
   ds_range       tp_cell_range,
   flds_to_cache  tp_col_agg_fns,  -- dynamically built from `pivot_axes` on the Pivot Table
   cached_fields  tp_cache_fields, -- dynamically built, indexed by col-heading
   cf_order       tp_data_ix_ord,
   wb_rel         PLS_INTEGER
);
TYPE tp_pivot_caches IS TABLE OF tp_pivot_cache INDEX BY PLS_INTEGER;
TYPE tp_pivot_table IS RECORD (
   pivot_table_id PLS_INTEGER,
   pivot_name     VARCHAR2(200),
   cache_id       PLS_INTEGER,
   on_sheet       PLS_INTEGER,
   location_tl    tp_cell_loc,
   pivot_axes     tp_pivot_axes,
   json_table     json_object_t,
   pivot_height   PLS_INTEGER,
   pivot_width    PLS_INTEGER
);
TYPE tp_pivot_tables IS TABLE OF tp_pivot_table INDEX BY PLS_INTEGER;
TYPE tp_pivots_list  IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;

-----
-- image/drawing/picture types
TYPE tp_drawing IS RECORD (
   img_id      PLS_INTEGER,
   row         PLS_INTEGER,
   col         PLS_INTEGER,
   scale       NUMBER,
   name        VARCHAR2(100),
   title       VARCHAR2(100),
   description VARCHAR2(4000)
);
TYPE tp_drawings IS TABLE OF tp_drawing INDEX BY PLS_INTEGER;

-----
-- sheet type
--
TYPE tp_sheet IS RECORD (
   wb_rel       PLS_INTEGER,
   rows         tp_rows,
   widths       tp_widths,
   name         VARCHAR2(100),
   freeze_rows  PLS_INTEGER,
   freeze_cols  PLS_INTEGER,
   autofilters  tp_autofilters,
   hyperlinks   tp_hyperlinks,
   col_fmts     tp_col_fmts,
   row_fmts     tp_row_fmts,
   comments     tp_comments,
   mergecells   tp_mergecells,
   validations  tp_validations,
   tabcolor     VARCHAR2(8),
   fontId       PLS_INTEGER,
   pivots_list  tp_pivots_list,
   drawings     tp_drawings
);
TYPE tp_sheets IS TABLE OF tp_sheet INDEX BY PLS_INTEGER;

-----
-- workbook types
--
TYPE tp_formulas IS TABLE OF VARCHAR2(32767) INDEX BY PLS_INTEGER;
TYPE tp_numFmts IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(100);
TYPE tp_fill IS RECORD (
   patternType VARCHAR2(30),
   fgRGB VARCHAR2(8),
   bgRGB VARCHAR2(8)
);
TYPE tp_fills IS TABLE OF tp_fill INDEX BY PLS_INTEGER;
TYPE tp_cellXfs IS TABLE OF tp_xf_fmt INDEX BY PLS_INTEGER;
TYPE tp_font IS RECORD (
   name      VARCHAR2(100),
   family    PLS_INTEGER,
   fontsize  NUMBER,
   theme     PLS_INTEGER,
   RGB       VARCHAR2(8),
   underline BOOLEAN,
   italic    BOOLEAN,
   bold      BOOLEAN
);
TYPE tp_fonts IS TABLE OF tp_font INDEX BY PLS_INTEGER;
TYPE tp_border IS RECORD (
   top    VARCHAR2(17),
   bottom VARCHAR2(17),
   left   VARCHAR2(17),
   right  VARCHAR2(17)
);
TYPE tp_borders IS TABLE OF tp_border INDEX BY PLS_INTEGER;
TYPE tp_strings IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(32767 char);
TYPE tp_str_ind IS TABLE OF VARCHAR2(32767 char) INDEX BY PLS_INTEGER;
TYPE tp_defined_names IS TABLE OF tp_cell_range INDEX BY VARCHAR2(100);

TYPE tp_image IS RECORD (
   img_blob    BLOB,
   img_hash    RAW(128),
   extension   VARCHAR2(5),
   width       PLS_INTEGER,
   height      PLS_INTEGER
);
TYPE tp_images IS TABLE OF tp_image INDEX BY PLS_INTEGER;

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
   defined_names tp_defined_names,
   formulas      tp_formulas,
   fontId        PLS_INTEGER,
   pivot_caches  tp_pivot_caches,
   pivot_tables  tp_pivot_tables,
   images        tp_images
);


wb_                   tp_book;
g_useXf_              BOOLEAN := true;
g_addtxt2utf8blob_tmp VARCHAR2(32767);

TYPE xml_attr IS RECORD (
   key   VARCHAR2(200),
   val   VARCHAR2(2000)
);
TYPE xml_attrs_arr IS TABLE OF xml_attr INDEX BY PLS_INTEGER;
--TYPE xml_attrs_arr IS TABLE OF VARCHAR2(2000) INDEX BY VARCHAR2(200);



---------------------------------------
---------------------------------------
-- 
-- Exception handling
--
--
-- Raise_App_Error()
--   Written as a wrapper function to make it easier to enter your own code if
--   you'd like to add some logging functionality or whatnot.
--
PROCEDURE Raise_App_Error (
   err_text_ IN VARCHAR2,
   p1_       IN VARCHAR2 := null,
   p2_       IN VARCHAR2 := null,
   p3_       IN VARCHAR2 := null,
   p4_       IN VARCHAR2 := null,
   p5_       IN VARCHAR2 := null,
   p6_       IN VARCHAR2 := null,
   p7_       IN VARCHAR2 := null,
   p8_       IN VARCHAR2 := null,
   p9_       IN VARCHAR2 := null,
   p0_       IN VARCHAR2 := null,
   repl_nl_  IN BOOLEAN  := true )
IS
BEGIN
   Cbh_Utils_API.Raise_App_Error (
      err_text_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_
   );
END Raise_App_Error;


---------------------------------------
---------------------------------------
-- 
-- Function Definitions - value getters
--
--
FUNCTION Get_Cell_Xf (
   sheet_ IN PLS_INTEGER,
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER ) RETURN tp_Xf_fmt;
FUNCTION Get_Cell_Xff (
   sheet_ IN PLS_INTEGER,
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER ) RETURN tp_Xf_fmt;
FUNCTION Get_Shared_String_Ix (
   string_ IN VARCHAR2 ) RETURN PLS_INTEGER;
FUNCTION Date_To_Xl_Nr (
   date_ IN DATE ) RETURN NUMBER;
FUNCTION Get_Cell_Value_Raw (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sh_     IN PLS_INTEGER,
   ss_ref_ IN BOOLEAN := true ) RETURN VARCHAR2;


---------------------------------------
---------------------------------------
-- 
-- General Helper Functions
--
--
PROCEDURE Trace (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true,
   quiet_   IN BOOLEAN  := false,
   indent_  IN NUMBER   := 0 )
IS
   m_  CLOB := lpad (' ', indent_ * 3) || msg_;
BEGIN
   Cbh_Utils_API.Trace (m_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_, quiet_);
END Trace;
FUNCTION Rep (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true ) RETURN VARCHAR2
IS BEGIN
   RETURN Cbh_Utils_API.Rep (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_);
END Rep;


---------------------------------------
---------------------------------------
--
-- Excel helpers
--
--
FUNCTION Get_Guid RETURN VARCHAR2
IS
   guid_ VARCHAR2(50) := RawToHex(sys_guid());
BEGIN
   RETURN '{' ||
      substr (guid_, 1,  8) || '-' || substr (guid_, 9,  4) || '-' ||
      substr (guid_, 13, 4) || '-' || substr (guid_, 17, 4) || '-' ||
      substr (guid_, 21, 12) || '}';
END Get_Guid;


---------------------------------------
---------------------------------------
--
-- XML generators and helpers
--
--
FUNCTION Xml_Date_Time (
   dt_ IN DATE ) RETURN VARCHAR2
IS BEGIN
   RETURN replace (to_char(dt_, 'YYYY-MM-DD_HH24:MI:SS'),'_','T');
END Xml_Date_Time;

FUNCTION Xml_Number (
   num_ IN NUMBER,
   fm_  IN VARCHAR2 := null ) RETURN VARCHAR2
IS
   mask_ VARCHAR2(99) := nvl (fm_, 'fm99999999999999999999.99999');
BEGIN
   RETURN rtrim (to_char (num_, mask_), '.');
END Xml_Number;

PROCEDURE Attr (
   key_   IN VARCHAR2,
   val_   IN VARCHAR2,
   attrs_ IN OUT NOCOPY xml_attrs_arr )
IS
   next_ix_ PLS_INTEGER := attrs_.count + 1;
BEGIN
   attrs_(next_ix_) := xml_attr ( key => key_, val => val_);
END Attr;

PROCEDURE nAtr (
   key_   IN VARCHAR2,
   val_   IN VARCHAR2,
   attrs_ IN OUT NOCOPY xml_attrs_arr )
IS BEGIN
   attrs_.delete;
   Attr (key_, val_, attrs_);
END nAtr;

FUNCTION Make_Tag (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomElement
IS
   el_ dbms_XmlDom.DomElement;
   ix_ VARCHAR2(200) := attrs_.first;
BEGIN
   el_ := CASE
      WHEN ns_ IS NOT null THEN Dbms_XmlDom.createElement (doc_, tag_name_, ns_)
      ELSE Dbms_XmlDom.createElement (doc_, tag_name_)
   END;
   FOR ix_ IN 1 .. attrs_.count LOOP
      Dbms_XmlDom.setAttribute (el_, attrs_(ix_).key, attrs_(ix_).val);
   END LOOP;
   RETURN el_;
END Make_Tag;

FUNCTION Make_Node (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode
IS
   nd_ dbms_XmlDom.DomNode := Dbms_XmlDom.makeNode (Make_Tag (doc_, tag_name_, ns_, attrs_));
BEGIN
   IF ns_ IS NOT null THEN
      Dbms_XmlDom.setPrefix (nd_, ns_);
   END IF;
   RETURN nd_;
END Make_Node;

PROCEDURE Make_Node (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := xml_attrs_arr() )
IS
   throw_nd_ dbms_XmlDom.DomNode;
BEGIN
   throw_nd_ := Make_Node (doc_, tag_name_, ns_, attrs_);
END Make_Node;

FUNCTION Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   ns_        IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode
IS BEGIN
   RETURN Dbms_XmlDom.appendChild (append_to_, Make_Node(doc_,tag_name_,ns_,attrs_));
END Xml_Node;

FUNCTION Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode
IS BEGIN
   RETURN Xml_Node (doc_, append_to_, tag_name_, '', attrs_);
END Xml_Node;

PROCEDURE Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   ns_        IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() )
IS
   throw_nd_ dbms_XmlDom.DomNode;
BEGIN
   throw_nd_ := Dbms_XmlDom.appendChild (append_to_, Make_Node(doc_,tag_name_,ns_,attrs_));
END Xml_Node;

PROCEDURE Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() )
IS BEGIN
   Xml_Node (doc_, append_to_, tag_name_, '', attrs_);
END Xml_Node;

PROCEDURE Xml_Text_Node (
   doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_    IN dbms_XmlDom.DomNode,
   tag_name_     IN VARCHAR2,
   text_content_ IN VARCHAR2,
   ns_           IN VARCHAR2,
   attrs_        IN xml_attrs_arr := xml_attrs_arr() )
IS
   throw_nd_ dbms_XmlDom.DomNode;
BEGIN
   throw_nd_ := Dbms_XmlDom.appendChild (
      Dbms_XmlDom.appendChild (append_to_, Make_Node(doc_,tag_name_,ns_,attrs_)),
      Dbms_XmlDom.makeNode (Dbms_XmlDom.createTextNode (doc_, text_content_))
   );
END Xml_Text_Node;

PROCEDURE Xml_Text_Node (
   doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_    IN dbms_XmlDom.DomNode,
   tag_name_     IN VARCHAR2,
   text_content_ IN VARCHAR2,
   attrs_        IN xml_attrs_arr := xml_attrs_arr() )
IS BEGIN
   Xml_Text_Node (doc_, append_to_, tag_name_, text_content_, '', attrs_);
END Xml_Text_Node;

PROCEDURE Xml_Text_Node (
   doc_         IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_   IN dbms_XmlDom.DomNode,
   tag_name_    IN VARCHAR2,
   num_content_ IN NUMBER,
   decimals_    IN NUMBER        := 0,
   ns_          IN VARCHAR2      := '',
   attrs_       IN xml_attrs_arr := xml_attrs_arr() )
IS BEGIN
   Xml_Text_Node (
      doc_          => doc_,
      append_to_    => append_to_,
      tag_name_     => tag_name_,
      text_content_ => Xml_Number (num_content_, decimals_),
      ns_           => ns_,
      attrs_        => attrs_
   );
END Xml_Text_Node;


---------------------------------------
---------------------------------------
--
-- Finnishing functions
--
--
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

PROCEDURE addtxt2utf8blob (
   txt_  IN            VARCHAR2,
   blob_ IN OUT NOCOPY BLOB )
IS BEGIN
   g_addtxt2utf8blob_tmp := g_addtxt2utf8blob_tmp || txt_;
EXCEPTION
   WHEN value_error THEN
      addtxt2utf8blob_finish (blob_);
      g_addtxt2utf8blob_tmp := txt_;
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
      utl_raw.substr (raw_, pos_, len_), utl_raw.little_endian
   );
END Raw2Num;

FUNCTION Little_Endian (
   big_   NUMBER,
   bytes_ PLS_INTEGER := 4 ) RETURN RAW
IS BEGIN
   RETURN utl_raw.substr (
      utl_raw.cast_from_binary_integer (big_, utl_raw.little_endian), 1, bytes_
   );
END Little_Endian;

FUNCTION Blob2Num (
   blob_ BLOB,
   len_  INTEGER,
   pos_  INTEGER ) RETURN NUMBER
IS BEGIN
   RETURN utl_raw.cast_to_binary_integer (
      dbms_lob.substr (blob_, len_, pos_), utl_raw.little_endian
   );
END Blob2Num;

PROCEDURE Add1File (
   zipped_blob_ IN OUT BLOB,
   filename_    IN VARCHAR2,
   content_     IN BLOB )
IS
   now_        DATE := sysdate;
   blob_       BLOB;
   len_        INTEGER;
   clen_       INTEGER;
   crc32_      RAW(4) := hextoraw('00000000');
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
      dbms_lob.createtemporary (zipped_blob_, true);
   END IF;
   name_raw_ := Utl_i18n.String_To_Raw (filename_, 'AL32UTF8');
   Dbms_Lob.Append (
      zipped_blob_,
      Utl_Raw.Concat(
         LOCAL_FILE_HEADER_, -- Local file header signature
         hextoraw('1400'),   -- version 2.0
         CASE WHEN name_raw_ = Utl_i18n.String_To_Raw (filename_, 'US8PC437')
            THEN hextoraw('0000') -- no General purpose bits
            ELSE hextoraw('0008') -- set Language encoding flag (EFS)
         END, CASE WHEN compressed_
            THEN hextoraw('0800') -- deflate
            ELSE hextoraw('0000') -- stored
         END,
         Little_Endian (
            to_number(to_char (now_, 'ss'))/2 + to_number(to_char (now_, 'mi'))*32 +
            to_number(to_char (now_, 'hh24'))*2048, 2
         ), -- File last modification time
         Little_Endian (
            to_number(to_char(now_,'dd')) + to_number(to_char(now_,'mm'))*32 +
            (to_number(to_char(now_,'yyyy'))-1980)*512, 2
         ), -- File last modification date
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

PROCEDURE Add1Xml (
   excel_    IN OUT NOCOPY BLOB,
   filename_ IN VARCHAR2,
   xml_      IN CLOB )
IS
   xml_blob_     BLOB;
   dest_offset_  INTEGER := 1;
   src_offset_   INTEGER := 1;
   lang_context_ INTEGER := Dbms_Lob.DEFAULT_LANG_CTX;
   warning_      INTEGER;
BEGIN
   Dbms_Lob.CreateTemporary (xml_blob_, true);
   Dbms_Lob.ConvertToBlob (
      xml_blob_, xml_, Dbms_Lob.LobMaxSize, dest_offset_, src_offset_,
      nls_charset_id('AL32UTF8'), lang_context_, warning_
   );
   Add1File (excel_, filename_, xml_blob_);
   Dbms_Lob.freetemporary(xml_blob_);
END Add1Xml;

PROCEDURE Finish_Zip (
   zipped_blob_ IN OUT BLOB )
IS
   nr_             PLS_INTEGER := 0;
   offset_            INTEGER;
   offs_dir_header_ INTEGER;
   offs_end_header_ INTEGER;
   watermark_         RAW(200) := Utl_Raw.Cast_To_Raw (
      'Implementation by Anton Scheffer, ' || VERSION_
   );
BEGIN
   offs_dir_header_ := dbms_lob.getlength (zipped_blob_);
   offset_ := 1;
   WHILE Dbms_Lob.Substr(zipped_blob_, utl_raw.length(LOCAL_FILE_HEADER_), offset_) = LOCAL_FILE_HEADER_ LOOP
      nr_ := nr_ + 1;
      Dbms_Lob.Append (
         zipped_blob_,
         Utl_Raw.Concat (
            hextoraw('504B0102'),      -- Central directory file header signature
            hextoraw('1400'),          -- version 2.0
            dbms_lob.substr(zipped_blob_, 26, offset_+4),
            hextoraw('0000'),          -- File comment length
            hextoraw('0000'),          -- Disk number where file starts
            hextoraw('0000'),          -- Internal file attributes => 0000=binary-file; 0100(ascii)=text-file
            CASE
               WHEN Dbms_Lob.Substr (
                  zipped_blob_, 1, offset_+30+blob2num(zipped_blob_,2,offset_+26)-1
               ) IN (hextoraw('2F'), hextoraw('5C'))
               THEN
                  hextoraw('10000000') -- a directory/folder
               ELSE
                  hextoraw('2000B681') -- a file
            END,                       -- External file attributes
            little_endian(offset_-1),  -- Relative offset of local file header
            dbms_lob.substr(zipped_blob_, blob2num(zipped_blob_,2,offset_+26),offset_+30) -- File name
         )
      );
      offset_ := offset_ + 30 +
         blob2num (zipped_blob_, 4, offset_+18 ) + -- compressed size
         blob2num (zipped_blob_, 2, offset_+26 ) + -- File name length
         blob2num (zipped_blob_, 2, offset_+28 );  -- Extra field length
   END LOOP;
   offs_end_header_ := dbms_lob.getlength(zipped_blob_);
   Dbms_Lob.Append (
       zipped_blob_,
       Utl_Raw.Concat (
          END_OF_CENTRAL_DIRECTORY_,                           -- End of central directory signature
          hextoraw ('0000'),                                   -- Number of this disk
          hextoraw ('0000'),                                   -- Disk where central directory starts
          little_endian (nr_, 2),                              -- Number of central directory records on this disk
          little_endian (nr_, 2),                              -- Total number of central directory records
          little_endian (offs_end_header_ - offs_dir_header_), -- Size of central directory
          little_endian (offs_dir_header_),                    -- Offset of start of central directory, relative to start of archive
          little_endian (nvl(Utl_Raw.Length(watermark_),0),2), -- ZIP file comment length
          watermark_
       )
    );
END Finish_Zip;


---------------------------------------
-- Print_Range()
--   A debug facility to output the data in a range
--
PROCEDURE Print_Range (
   range_ IN tp_cell_range )
IS
BEGIN
   Trace (q'[
sheet_id     : :P1
top-left     : cell: (c: :P2, r: :P3)
bottom-right : cell: (c: :P4, r: :P5)
defined-name : :P6
   ]',
      to_char(range_.sheet_id), to_char(range_.tl.c), to_char(range_.tl.r),
      to_char(range_.br.c), to_char(range_.br.r), range_.defined_name
   );
END Print_Range;

---------------------------------------
---------------------------------------
--
-- Cell reference converters
-- > Alfanumeric to number reference.  Useful as a helper for generating Excel
--   formulas such as `sum(A3:X3)`.  But also required when building XML parts
-- > Alfan_Col() => helps to convert (2, 3) => B3
-- > Col_Alfan() => helps to convert B3 => (2, 3)
--
--
FUNCTION Col_Alfan(
   col_ IN VARCHAR2 ) RETURN PLS_INTEGER
IS BEGIN
   RETURN ascii(substr(col_,-1)) - 64
      + nvl((ascii(substr(col_,-2,1))-64) * 26, 0)
      + nvl((ascii(substr(col_,-3,1))-64) * 676, 0);
END Col_Alfan;

FUNCTION Alfan_Col (
   col_ IN PLS_INTEGER ) RETURN VARCHAR2
IS BEGIN
   RETURN CASE
      WHEN col_ > 702 THEN chr(64+trunc((col_-27)/676)) || chr(65+mod(trunc((col_-1)/26)-1, 26)) || chr(65+mod(col_-1, 26))
      WHEN col_ > 26  THEN chr(64+trunc((col_-1)/26)) || chr(65+mod(col_-1, 26))
      ELSE chr(64+col_)
   END;
END Alfan_Col;

FUNCTION Alfan_Cell (
   col_  IN PLS_INTEGER,
   row_  IN PLS_INTEGER,
   fix1_ IN BOOLEAN := false,
   fix2_ IN BOOLEAN := false ) RETURN VARCHAR2
IS
   d1_  VARCHAR2(1) := CASE WHEN fix1_ THEN '$' END;
   d2_  VARCHAR2(1) := CASE WHEN fix2_ THEN '$' END;
BEGIN
   RETURN d1_ || Alfan_Col (col_) || d2_ || to_char(row_);
END Alfan_Cell;

FUNCTION Alfan_Cell (
   loc_ IN OUT NOCOPY tp_cell_loc ) RETURN VARCHAR2
IS BEGIN
   RETURN Alfan_Cell (loc_.c, loc_.r, loc_.fixc, loc_.fixr);
END Alfan_Cell;

PROCEDURE Alfan_Cell (
   loc_ IN OUT NOCOPY tp_cell_loc )
IS
   throw_ VARCHAR2(12);
BEGIN
   throw_ := Alfan_Cell (loc_.c, loc_.r, loc_.fixc, loc_.fixr);
END Alfan_Cell;

FUNCTION Alfan_Range (
   col_tl_  IN PLS_INTEGER,
   row_tl_  IN PLS_INTEGER,
   col_br_  IN PLS_INTEGER,
   row_br_  IN PLS_INTEGER,
   fix_tlc_ IN BOOLEAN := false,
   fix_tlr_ IN BOOLEAN := false,
   fix_brc_ IN BOOLEAN := false,
   fix_brr_ IN BOOLEAN := false ) RETURN VARCHAR2
IS BEGIN
   IF col_tl_ IS null OR row_tl_ IS null OR col_br_ IS null OR row_br_ IS null THEN
      RETURN 'A1';
   END IF;
   RETURN Alfan_Cell (col_tl_, row_tl_, fix_tlc_, fix_tlr_) || ':' ||
          Alfan_Cell (col_br_, row_br_, fix_brc_, fix_brr_);
END Alfan_Range;

FUNCTION Alfan_Range (
   range_ IN OUT NOCOPY tp_cell_range ) RETURN VARCHAR2
IS BEGIN
   RETURN Alfan_Range (
      range_.tl.c, range_.tl.r, range_.br.c, range_.br.r,
      range_.tl.fixc, range_.tl.fixr, range_.br.fixc, range_.br.fixr
   );
END Alfan_Range;

FUNCTION Alfan_Sheet_Range (
   sheet_name_ IN VARCHAR2,
   col_tl_     IN PLS_INTEGER,
   row_tl_     IN PLS_INTEGER,
   col_br_     IN PLS_INTEGER,
   row_br_     IN PLS_INTEGER,
   fix_tlc_    IN BOOLEAN := true,
   fix_tlr_    IN BOOLEAN := true,
   fix_brc_    IN BOOLEAN := true,
   fix_brr_    IN BOOLEAN := true ) RETURN VARCHAR2
IS
   sheet_prefix_ VARCHAR2(103) := CASE WHEN sheet_name_ IS NOT null THEN
      '''' || sheet_name_ || '''!'
   END;
BEGIN
   RETURN sheet_prefix_ || Alfan_Range (
      col_tl_, row_tl_, col_br_, row_br_, fix_tlc_, fix_tlr_, fix_brc_, fix_brr_
   );
END Alfan_Sheet_Range;

FUNCTION Alfan_Sheet_Range (
   sheet_   IN PLS_INTEGER,
   col_tl_  IN PLS_INTEGER,
   row_tl_  IN PLS_INTEGER,
   col_br_  IN PLS_INTEGER,
   row_br_  IN PLS_INTEGER,
   fix_tlc_ IN BOOLEAN := true,
   fix_tlr_ IN BOOLEAN := true,
   fix_brc_ IN BOOLEAN := true,
   fix_brr_ IN BOOLEAN := true ) RETURN VARCHAR2
IS BEGIN
   RETURN Alfan_Sheet_Range (
      wb_.sheets(sheet_).name, col_tl_, row_tl_, col_br_, row_br_,
      fix_tlc_, fix_tlr_, fix_brc_, fix_brr_
   );
END Alfan_Sheet_Range;

FUNCTION Alfan_Sheet_Range (
   range_ IN tp_cell_range ) RETURN VARCHAR2
IS BEGIN
   RETURN Alfan_Sheet_Range (
      range_.sheet_id, range_.tl.c, range_.tl.r, range_.br.c, range_.br.r,
      range_.tl.fixc, range_.tl.fixr, range_.br.fixc, range_.br.fixr
   );
END Alfan_Sheet_Range;

FUNCTION Sheet_Name (
   sheet_ IN PLS_INTEGER ) RETURN VARCHAR2
IS BEGIN
   RETURN wb_.sheets(sheet_).name;
END Sheet_Name;

FUNCTION Sheet_Name (
   range_ IN tp_cell_range ) RETURN VARCHAR2
IS BEGIN
   RETURN wb_.sheets(range_.sheet_id).name;
END Sheet_Name;

FUNCTION Range_Height (
   range_       IN tp_cell_range,
   include_hdr_ IN BOOLEAN := false ) RETURN PLS_INTEGER
IS
   add_hdr_ PLS_INTEGER := CASE WHEN include_hdr_ THEN 1 ELSE 0 END;
BEGIN
   RETURN range_.br.r - range_.tl.r + add_hdr_;
END Range_Height;

FUNCTION Range_Width (
   range_ IN tp_cell_range ) RETURN PLS_INTEGER
IS BEGIN
   RETURN range_.br.c - range_.tl.c + 1;
END Range_Width;

PROCEDURE Add_Col_Headings_To_Range (
   range_ IN OUT NOCOPY tp_cell_range,
   sheet_ IN PLS_INTEGER := null )
IS
   i_   PLS_INTEGER := 1;
   row_ PLS_INTEGER := range_.tl.r;
   sh_  PLS_INTEGER := coalesce (range_.sheet_id, sheet_, wb_.sheets.count);
BEGIN
   IF range_.col_names.count = 0 THEN -- else assume `col_names` is correctly filled out
      FOR c_ IN range_.tl.c .. range_.br.c LOOP
         range_.col_names(i_) := wb_.sheets(sh_).rows(row_)(c_).ora_value.str_val;
         i_ := i_ + 1;
      END LOOP;
   END IF;
END Add_Col_Headings_To_Range;

FUNCTION Range_Col_Head_Name (
   range_    IN tp_cell_range,
   col_offs_ IN PLS_INTEGER,
   sheet_    IN PLS_INTEGER := null ) RETURN VARCHAR2
IS
   col_ PLS_INTEGER := range_.tl.c + col_offs_ - 1;
   row_ PLS_INTEGER := range_.tl.r;
   sh_  PLS_INTEGER := coalesce (range_.sheet_id, sheet_, wb_.sheets.count);
BEGIN
   IF range_.col_names.exists(col_offs_) THEN
      RETURN range_.col_names(col_offs_);
   ELSE
      RETURN wb_.sheets(sh_).rows(row_)(col_).ora_value.str_val; -- assume the column headre is a chr-value
   END IF;
END Range_Col_Head_Name;

FUNCTION Range_Col_NumFmtId (
   range_    IN tp_cell_range,
   col_offs_ IN PLS_INTEGER,
   sheet_    IN PLS_INTEGER := null ) RETURN PLS_INTEGER
IS
   col_ PLS_INTEGER := range_.tl.c + col_offs_ - 1;
   row_ PLS_INTEGER := range_.tl.r + 1;
   sh_  PLS_INTEGER := coalesce (range_.sheet_id, sheet_, wb_.sheets.count);
BEGIN
   RETURN Get_Cell_Xff (sh_, col_, row_).numFmtId;
END Range_Col_NumFmtId;

FUNCTION Range_Unique_Data_Ord (
   range_    IN tp_cell_range,
   col_offs_ IN PLS_INTEGER,
   sheet_    IN PLS_INTEGER := null ) RETURN tp_data_ix_ord
IS
   col_       PLS_INTEGER := range_.tl.c + col_offs_ - 1;
   row_start_ PLS_INTEGER := range_.tl.r + 1; -- allow for header row
   row_end_   PLS_INTEGER := range_.br.r;
   sh_        PLS_INTEGER := coalesce (range_.sheet_id, sheet_, wb_.sheets.count);
   val_       VARCHAR2(2000);
   unq_data_  tp_unique_data;
   ord_data_  tp_data_ix_ord;
   new_ix_    PLS_INTEGER := 0; -- pivotCacheRecord uses a base of zero
BEGIN
   IF sh_ IS null THEN
      Raise_App_Error ('A dataset range must have a Sheet Id, in Range_Unique_Data_Ord()');
   END IF;
   FOR r_ IN row_start_ .. row_end_ LOOP
      val_ := Get_Cell_Value_Raw (col_, r_, sh_, false);
      IF not unq_data_.exists(val_) THEN
         unq_data_(val_)    := new_ix_;
         ord_data_(new_ix_) := val_;
         new_ix_ := new_ix_ + 1;
      END IF;
   END LOOP;
   RETURN ord_data_;
END Range_Unique_Data_Ord;

FUNCTION Ord_Data_To_Unique (
   ixs_ordered_ IN tp_data_ix_ord ) RETURN tp_unique_data
IS
   data_uq_ tp_unique_data;
BEGIN
   FOR ix_ IN ixs_ordered_.first .. ixs_ordered_.last LOOP
      data_uq_(ixs_ordered_(ix_)) := ix_;
   END LOOP;
   RETURN data_uq_;
END Ord_Data_To_Unique;

PROCEDURE Range_Col_Min_Max_Values (
   range_    IN     tp_cell_range,
   col_offs_ IN     PLS_INTEGER,
   min_val_  IN OUT NOCOPY NUMBER,
   max_val_  IN OUT NOCOPY NUMBER,
   sheet_    IN     PLS_INTEGER := null )
IS
   col_       PLS_INTEGER := range_.tl.c + col_offs_ - 1;
   row_start_ PLS_INTEGER := range_.tl.r + 1; -- allow for header row
   row_end_   PLS_INTEGER := range_.br.r;
   sh_        PLS_INTEGER := coalesce (range_.sheet_id, sheet_, wb_.sheets.count);
   val_       NUMBER;
BEGIN
   min_val_ := null;
   max_val_ := null;
   FOR r_ IN row_start_ .. row_end_ LOOP
      val_ := wb_.sheets(sh_).rows(r_)(col_).ora_value.num_val;
      min_val_ := CASE
         WHEN min_val_ IS null THEN val_
         WHEN min_val_ > val_  THEN val_
         ELSE min_val_
      END;
      max_val_ := CASE
         WHEN max_val_ IS null THEN val_
         WHEN max_val_ < val_  THEN val_
         ELSE max_val_
      END;
   END LOOP;
END Range_Col_Min_Max_Values;


---------------------------------------
---------------------------------------
-- 
-- Cell value getters
--
--
FUNCTION Get_Cell_Value_Num (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null ) RETURN NUMBER
IS BEGIN
   RETURN wb_.sheets(sheet_).rows(row_)(col_).ora_value.num_val;
END Get_Cell_Value_Num;

FUNCTION Get_Cell_Value_Str (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null ) RETURN VARCHAR2
IS BEGIN
   RETURN wb_.sheets(sheet_).rows(row_)(col_).ora_value.str_val;
END Get_Cell_Value_Str;

FUNCTION Get_Cell_Value_Date (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER ) RETURN DATE
IS BEGIN
   RETURN wb_.sheets(sheet_).rows(row_)(col_).ora_value.dt_val;
END Get_Cell_Value_Date;

-----
-- Get_Cell_Value_Raw()
--   Getting the raw value means different things depending on the datatype of
--   the cell:
--     -> for strongs, we (optionally) fetch the shared-string reference, else
--        the string value itself
--     -> for numbers, we fetch the raw value, no thousand-separated, currency
--        or other visual beautification
--     -> for dates, we return a number according to Excel's base-1900 system
--
FUNCTION Get_Cell_Value_Raw (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sh_     IN PLS_INTEGER,
   ss_ref_ IN BOOLEAN := true ) RETURN VARCHAR2
IS
   ss_id_ PLS_INTEGER := -1;
BEGIN
   IF wb_.sheets(sh_).rows(row_)(col_).datatype = CELL_DT_STRING_ THEN
      IF ss_ref_ THEN
         ss_id_ := Get_Shared_String_Ix (Get_Cell_Value_Str(col_,row_,sh_));
      END IF;
      RETURN CASE WHEN ss_id_ = -1 THEN Get_Cell_Value_Str(col_,row_,sh_) ELSE to_char(ss_id_) END;
   ELSIF wb_.sheets(sh_).rows(row_)(col_).datatype = CELL_DT_NUMBER_ THEN
      RETURN to_char(Get_Cell_Value_Num (col_, row_, sh_));
   ELSIF wb_.sheets(sh_).rows(row_)(col_).datatype = CELL_DT_DATE_ THEN
      RETURN to_char (Date_To_Xl_Nr(Get_Cell_Value_Date (col_, row_, sh_)));
   END IF;
END Get_Cell_Value_Raw;

FUNCTION Get_Cell_Cache_Value (
   col_          IN PLS_INTEGER,
   row_          IN PLS_INTEGER,
   sheet_        IN PLS_INTEGER,
   shared_items_ IN tp_unique_data ) RETURN VARCHAR2
IS
   value_ VARCHAR2(32000) := Get_Cell_Value_Raw (col_, row_, sheet_, false);
BEGIN
   RETURN CASE
      WHEN not shared_items_.exists(value_) THEN value_
      ELSE to_char(shared_items_(value_))
   END;
END Get_Cell_Cache_Value;

-----
-- Get_Cell_Cache_Tag()
--   Useful in the pivotCacheRecords.xml files
--
FUNCTION Get_Cell_Cache_Tag (
   col_      IN PLS_INTEGER,
   row_      IN PLS_INTEGER,
   sheet_    IN PLS_INTEGER,
   agg_type_ IN VARCHAR2 ) RETURN VARCHAR2
IS BEGIN
   IF agg_type_ IN ('col','row','filter') THEN
      RETURN 'x';
   ELSIF wb_.sheets(sheet_).rows(row_)(col_).datatype = CELL_DT_STRING_ THEN
      RETURN 's';
   ELSIF wb_.sheets(sheet_).rows(row_)(col_).datatype = CELL_DT_NUMBER_ THEN
      RETURN 'n';
   ELSIF wb_.sheets(sheet_).rows(row_)(col_).datatype = CELL_DT_DATE_ THEN
      RETURN 'n';
   END IF;
END Get_Cell_Cache_Tag;

-----
-- Get_Cell_Value_Fmt()
--   One can imagine a time when we'd need the cell's value in the format that
--   the user desires to see it.  However, I haven't yet found a use case that
--   matches my imagination yet, hence I've commented this out.
--
/*FUNCTION Get_Cell_Value_Fmt (
   col_     IN PLS_INTEGER,
   row_     IN PLS_INTEGER,
   sheet_   IN PLS_INTEGER,
   num_fmt_ IN VARCHAR2 := null,
   ss_ref_  IN BOOLEAN  := true ) RETURN VARCHAR2 -- ss = shared string
IS
   datatype_ VARCHAR2(30) := wb_.sheets(sheet_).rows(row_)(col_).datatype;
   fm_       VARCHAR2(100); -- foramt-mask
   ss_id_    PLS_INTEGER  := -1;
   ret_str_  VARCHAR2(32000);
BEGIN
   CASE wb_.sheets(sheet_).rows(row_)(col_).datatype

      WHEN CELL_DT_STRING_ THEN
         IF ss_ref_ THEN
           ss_id_ := Get_Shared_String_Ix (wb_.sheets(sheet_).rows(row_)(col_).ora_value.str_val);
         END IF;
         ret_str_ := CASE WHEN ss_ = -1 THEN Get_Cell_Value_Str(col_,row_,sheet_) ELSE to_char(ss_) END;

      WHEN CELL_DT_NUMBER_ THEN
         ret_str_ := CASE
            WHEN num_fmt_ IS null THEN to_char (Get_Cell_Value_Num(col_,row_,sheet_))
            ELSE to_char (Get_Cell_Value_Num(col_,row_,sheet_), num_fmt_)
         END;

      WHEN CELL_DT_DATE_ THEN
         fm_ := nvl (num_fmt_, 'YYYY-MM-DD-HH24:MI');
         ret_str_ := to_char (Get_Cell_Value_Date (col_, row_, sheet_), fm_);

   END CASE;
   RETURN ret_str_;
END Get_Cell_Value_Fmt;*/

---------------------------------------
---------------------------------------
--
-- Functions that build the internal PL/SQL model of the Excel sheet
--
--
PROCEDURE Clear_Workbook
IS
   s_      PLS_INTEGER := wb_.sheets.first;
   row_ix_ PLS_INTEGER;
BEGIN
   WHILE s_ IS NOT null LOOP
      row_ix_ := wb_.sheets(s_).rows.first;
      WHILE row_ix_ IS NOT null LOOP
         wb_.sheets(s_).rows(row_ix_).delete();
         row_ix_ := wb_.sheets(s_).rows.next(row_ix_);
      END LOOP;
      wb_.sheets(s_).rows.delete();
      wb_.sheets(s_).widths.delete();
      wb_.sheets(s_).autofilters.delete();
      wb_.sheets(s_).hyperlinks.delete();
      wb_.sheets(s_).col_fmts.delete();
      wb_.sheets(s_).row_fmts.delete();
      wb_.sheets(s_).comments.delete();
      wb_.sheets(s_).mergecells.delete();
      wb_.sheets(s_).validations.delete();
      wb_.sheets(s_).drawings.delete();
      s_ := wb_.sheets.next(s_);
   END LOOP;
   wb_.strings.delete();
   wb_.str_ind.delete();
   wb_.fonts.delete();
   wb_.fills.delete();
   wb_.borders.delete();
   wb_.numFmts.delete();
   wb_.cellXfs.delete();
   wb_.defined_names.delete();
   wb_.formulas.delete();
   FOR i_ IN 1 .. wb_.images.count LOOP
      dbms_lob.freeTemporary (wb_.images(i_).img_blob);
   END LOOP;
   wb_.images.delete();
   wb_ := null;
END Clear_Workbook;

PROCEDURE Set_Tabcolor (
   tabcolor_ VARCHAR2,
   sheet_    PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).tabcolor := substr(tabcolor_, 1, 8);
END Set_Tabcolor;

FUNCTION New_Sheet (
   sheetname_ VARCHAR2 := null,
   tab_color_ VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   s_ PLS_INTEGER := wb_.sheets.count + 1;
BEGIN
   wb_.sheets(s_).name := nvl (
      Dbms_XmlGen.Convert(translate(sheetname_, 'a/\[]*:?', 'a')),
      'Sheet' || s_
   );
   IF wb_.strings.count = 0 THEN
      wb_.str_cnt := 0;
   END IF;
   IF wb_.fonts.count = 0 THEN
      wb_.fontid := Get_Font('Calibri');
   END IF;
   IF wb_.fills.count = 0 THEN
      Get_Fill('none');
      Get_Fill('gray125');
   END IF;
   IF wb_.borders.count = 0 THEN
      Get_Border ('', '', '', '');
   END IF;
   Set_TabColor(tab_color_, s_);
   wb_.sheets(s_).fontId := wb_.fontId;
   RETURN s_;
END New_Sheet;

PROCEDURE New_Sheet (
   sheetname_ VARCHAR2 := null,
   tab_color_ VARCHAR2 := null )
IS
   throw_ PLS_INTEGER;
BEGIN
   throw_ := New_Sheet (sheetname_, tab_color_); --ignore
END New_Sheet;

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 )
IS BEGIN
   wb_.sheets(sheet_).name := nvl (
      Dbms_xmlgen.Convert (translate(name_, 'a/\[]*:?', 'a')),
      'Sheet'  || sheet_
   );
END Set_Sheet_Name;

-----
-- Set_Col_Width_By_Format()
-- Set_Column_Width()
--   These two functions have the same effect on the resulting Excel document,
--   but that the Set_Col_Width_By_Format() version is smarter, using a format
--   we give it to calculate the necessary width.  This is useful for currency
--   formats for example.
--   For the other, the value we pass represents the number of characters that
--   we'd like to see in a column.  It assumes a Calibri font, size 11.
PROCEDURE Set_Col_Width_By_Format (
   sheet_  IN PLS_INTEGER,
   col_    IN PLS_INTEGER,
   format_ IN VARCHAR2 )
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
   IF wb_.sheets(sheet_).widths.exists(col_) THEN
      wb_.sheets(sheet_).widths(col_) := greatest(
         wb_.sheets(sheet_).widths(col_), width_
      );
   ELSE
      wb_.sheets(sheet_).widths(col_) := greatest(width_, 8.43);
   END IF;
END Set_Col_Width_By_Format;

PROCEDURE Set_Column_Width (
   col_   PLS_INTEGER,
   width_ NUMBER,
   sheet_ PLS_INTEGER := null )
IS
   w_  NUMBER      := trunc(round(width_*7)*256/7)/256;
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).widths(col_) := w_;
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
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).col_fmts(col_).numFmtId  := numFmtId_;
   wb_.sheets(sh_).col_fmts(col_).fontId    := fontId_;
   wb_.sheets(sh_).col_fmts(col_).fillId    := fillId_;
   wb_.sheets(sh_).col_fmts(col_).borderId  := borderId_;
   wb_.sheets(sh_).col_fmts(col_).alignment := alignment_;
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
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
   c_  tp_cells;
BEGIN
   wb_.sheets(sh_).row_fmts(row_).numFmtId  := numFmtId_;
   wb_.sheets(sh_).row_fmts(row_).fontId    := fontId_;
   wb_.sheets(sh_).row_fmts(row_).fillId    := fillId_;
   wb_.sheets(sh_).row_fmts(row_).borderId  := borderId_;
   wb_.sheets(sh_).row_fmts(row_).alignment := alignment_;
   wb_.sheets(sh_).row_fmts(row_).height    := trunc(height_*4/3)*3/4;
   IF not wb_.sheets(sh_).rows.exists(row_) THEN
      wb_.sheets(sh_).rows(row_) := c_;
   END IF;
END Set_Row;


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
   format_mask_ VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   fmt_id_ PLS_INTEGER;
BEGIN
   IF format_mask_ IS null THEN
      fmt_id_ := 0;
   ELSIF wb_.numFmts.exists(format_mask_) THEN
      fmt_id_ := wb_.numFmts(format_mask_);
   ELSE
      fmt_id_ := wb_.numFmts.count + 164;
      wb_.numFmts(format_mask_) := fmt_id_;
   END IF;
   RETURN fmt_id_;
END Get_NumFmt;

FUNCTION Get_Num_Format_Mask (
   num_fmt_id_ IN PLS_INTEGER ) RETURN VARCHAR2
IS
   fmt_mask_ VARCHAR2(100) := wb_.numFmts.first;
BEGIN
   IF num_fmt_id_ = 0 THEN
      RETURN '';
   END IF;
   WHILE fmt_mask_ IS NOT null LOOP
      EXIT WHEN wb_.numFmts(fmt_mask_) = num_fmt_id_;
      fmt_mask_ := wb_.numFmts.next(fmt_mask_);
   END LOOP;
   RETURN fmt_mask_;
END Get_Num_Format_Mask;

PROCEDURE Add_NumFmt (
   fmt_id_ IN VARCHAR2,
   format_ IN VARCHAR2 )
IS BEGIN
   numFmt_(fmt_id_) := format_;
END Add_NumFmt;

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
      wb_.fontId := ix_;
   ELSE
      wb_.sheets(sheet_).fontId := ix_;
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
   IF wb_.fonts.count > 0 THEN
      FOR f_ IN 0 .. wb_.fonts.count - 1 LOOP
         IF (     wb_.fonts(f_).name      = name_
              AND wb_.fonts(f_).family    = family_
              AND wb_.fonts(f_).fontsize  = fontsize_
              AND wb_.fonts(f_).theme     = theme_
              AND wb_.fonts(f_).underline = underline_
              AND wb_.fonts(f_).italic    = italic_
              AND wb_.fonts(f_).bold      = bold_
              AND (     wb_.fonts(f_).rgb = rgb_
                    OR (wb_.fonts(f_).rgb IS null AND rgb_ IS null)
              )
         ) THEN
            RETURN f_;
         END IF;
      END LOOP;
   END IF;
   ix_ := wb_.fonts.count;
   wb_.fonts(ix_).name      := name_;
   wb_.fonts(ix_).family    := family_;
   wb_.fonts(ix_).fontsize  := fontsize_;
   wb_.fonts(ix_).theme     := theme_;
   wb_.fonts(ix_).underline := underline_;
   wb_.fonts(ix_).italic    := italic_;
   wb_.fonts(ix_).bold      := bold_;
   wb_.fonts(ix_).rgb       := rgb_;
   RETURN ix_;
END Get_Font;


FUNCTION Get_Fill (
   patternType_ VARCHAR2,
   fgRGB_       VARCHAR2 := null,
   bgRGB_       VARCHAR2 := null ) RETURN PLS_INTEGER
IS
   ix_ PLS_INTEGER;
BEGIN
   IF wb_.fills.count > 0 THEN
      FOR f_ IN 0 .. wb_.fills.count - 1 LOOP
         IF (   wb_.fills(f_).patternType = patternType_
            AND nvl(wb_.fills(f_).fgRGB, 'x') = nvl(upper(fgRGB_), 'x')
            AND nvl(wb_.fills(f_).bgRGB, 'x') = nvl(upper(bgRGB_), 'x')
         ) THEN
            RETURN f_;
         END IF;
      END LOOP;
   END IF;
   ix_ := wb_.fills.count;
   wb_.fills(ix_).patternType := patternType_;
   wb_.fills(ix_).fgRGB       := upper(fgRGB_);
   wb_.fills(ix_).bgRGB       := upper(bgRGB_);
   RETURN ix_;
END Get_Fill;

PROCEDURE Get_Fill (
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,
   bgRGB_       IN VARCHAR2 := null )
IS
   throw_ PLS_INTEGER;
BEGIN
   throw_ := Get_Fill (patternType_, fgRGB_, bgRGB_); --ignore
END Get_Fill;

PROCEDURE Add_Fill (
   fill_id_     IN VARCHAR2,
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,
   bgRGB_       IN VARCHAR2 := null )
IS BEGIN
   fills_(fill_id_) := Get_Fill (patternType_, fgRGB_, bgRGB_);
END Add_Fill;


FUNCTION Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' ) RETURN PLS_INTEGER
IS
   ix_ PLS_INTEGER;
BEGIN
   IF wb_.borders.count > 0 THEN
      FOR b_ IN 0 .. wb_.borders.count - 1 LOOP
         IF (   nvl(wb_.borders(b_).top,    'x') = nvl(top_, 'x')
            AND nvl(wb_.borders(b_).bottom, 'x') = nvl(bottom_, 'x')
            AND nvl(wb_.borders(b_).left,   'x') = nvl(left_, 'x')
            AND nvl(wb_.borders(b_).right,  'x') = nvl(right_, 'x')
         ) THEN
            RETURN b_;
         END IF;
      END LOOP;
   END IF;
   ix_ := wb_.borders.count;
   wb_.borders(ix_).top    := top_;
   wb_.borders(ix_).bottom := bottom_;
   wb_.borders(ix_).left   := left_;
   wb_.borders(ix_).right  := right_;
   RETURN ix_;
END Get_Border;

PROCEDURE Get_Border (
   top_    IN VARCHAR2 := 'thin',
   bottom_ IN VARCHAR2 := 'thin',
   left_   IN VARCHAR2 := 'thin',
   right_  IN VARCHAR2 := 'thin' )
IS
   throw_ NUMBER;
BEGIN
   throw_ := Get_Border (top_, bottom_, left_, right_); -- ignore
END Get_Border;

-----
-- Add_Border_To_Cell()
--   This function applies a border to a given cell while also preserving that
--   cell's existing styles.  Note that if we ONLY want to apply our border to
--   the right-wall of the cell, and preserve the border-styles of the other 3
--   walls, then we should leave those other 3 values null.  If you explicitly
--   need to unset a border, you can pass in the value 'none'
--
PROCEDURE Add_Border_To_Cell (
   col_     IN PLS_INTEGER,
   row_     IN PLS_INTEGER,
   top_     IN VARCHAR2    := '',
   bottom_  IN VARCHAR2    := '',
   left_    IN VARCHAR2    := '',
   right_   IN VARCHAR2    := '',
   sheet_   IN PLS_INTEGER := null )
IS
   sh_          PLS_INTEGER  := nvl(sheet_, wb_.sheets.count);
   Xf_          tp_Xf_fmt    := Get_Cell_Xff(sh_, col_, row_);
   cell_border_ tp_border    := wb_.borders(Xf_.borderId);
   cell_dt_     VARCHAR2(30) := wb_.sheets(sh_).rows(row_)(col_).datatype;
   border_id_   PLS_INTEGER;
BEGIN

   cell_border_.top    := nvl (top_,    cell_border_.top);
   cell_border_.bottom := nvl (bottom_, cell_border_.bottom);
   cell_border_.left   := nvl (left_,   cell_border_.left);
   cell_border_.right  := nvl (right_,  cell_border_.right);
   border_id_          := Get_Border (
      cell_border_.top, cell_border_.bottom, cell_border_.left, cell_border_.right
   );

   IF cell_dt_ = CELL_DT_NUMBER_ THEN
      Cell (
         col_, row_, Get_Cell_Value_Num (col_, row_, sh_), --wb_.sheets(sh_).rows(row_)(col_).ora_value.num_val,
         Xf_.numFmtId, Xf_.fontId, Xf_.fillId, border_id_, Xf_.alignment, sh_
      );
   ELSIF cell_dt_ = CELL_DT_STRING_ THEN
      Cell (
         col_, row_, Get_Cell_Value_Str (col_, row_, sh_), --wb_.sheets(sh_).rows(row_)(col_).ora_value.str_val,
         Xf_.numFmtId, Xf_.fontId, Xf_.fillId, border_id_, Xf_.alignment, sh_
      );
   ELSIF cell_dt_ = CELL_DT_DATE_ THEN
      Cell (
         col_, row_, Get_Cell_Value_Date (col_, row_, sh_), --wb_.sheets(sh_).rows(row_)(col_).ora_value.dt_val,
         Xf_.numFmtId, Xf_.fontId, Xf_.fillId, border_id_, Xf_.alignment, sh_
      );
   END IF;

END Add_Border_To_Cell;

-----
-- Add_Border_To_Range()
--   Take a range of cells and put a border around it!  The procedure will not
--   override other settings in that that range of cells even if some of those
--   other settings have set borders on some of the internal cells.
--   The parameters of this function need to be changed to accept tl/br combos
--   rather than height and width, in order for it to be consistent with other
--   range management functions.
--
PROCEDURE Add_Border_To_Range (
   cell_left_ IN PLS_INTEGER,
   cell_top_  IN PLS_INTEGER,
   width_     IN PLS_INTEGER,
   height_    IN PLS_INTEGER,
   style_     IN VARCHAR2    := 'medium', -- thin|medium|thick|dotted...
   sheet_     IN PLS_INTEGER := null )
IS
   sh_         PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
   col_start_  PLS_INTEGER := cell_left_;
   col_end_    PLS_INTEGER := cell_left_ + width_ - 1;
   row_start_  PLS_INTEGER := cell_top_;
   row_end_    PLS_INTEGER := cell_top_ + height_ - 1;
BEGIN

   -- first we should catch any invalid parameter combinations
   IF width_ < 1 OR height_ < 1 THEN
      Raise_App_Error ('Width and height of a border-range must be greater than zero');

   -- for a 1 x 1 span...
   ELSIF width_ = 1 AND height_ = 1 THEN
      Add_Border_To_Cell (cell_left_, cell_top_, style_, style_, style_, style_, sh_);

   -- for a n x 1 span...
   ELSIF height_ = 1 THEN
      Add_Border_To_Cell (cell_left_, cell_top_, style_, style_, style_, '', sh_);
      FOR col_ IN (cell_left_+1) .. (cell_left_+width_-2) LOOP
         Add_Border_To_Cell (col_, cell_top_, style_, style_, '', '', sh_);
      END LOOP;
      Add_Border_To_Cell (cell_left_+width_-1, cell_top_, style_, style_, '', style_, sh_);

   -- for a 1 x n span
   ELSIF width_ = 1 THEN
      Add_Border_To_Cell (cell_left_, cell_top_, style_, '', style_, style_, sh_);
      FOR row_ IN (cell_top_+1) .. (cell_top_+height_-2) LOOP
         Add_Border_To_Cell (cell_left_, row_, '', '', style_, style_, sh_);
      END LOOP;
      Add_Border_To_Cell (cell_left_, cell_top_+height_-1, '', style_, style_, style_, sh_);

   -- for an n x m span
   ELSE

      FOR col_ IN col_start_ .. col_end_ LOOP
         FOR row_ IN row_start_ .. row_end_ LOOP

            IF col_ = col_start_ THEN -- first column
               IF row_ = row_start_ THEN
                  Add_Border_To_Cell (col_, row_, style_, '', style_, '', sh_); -- top-left
               ELSIF row_ = row_end_ THEN
                  Add_Border_To_Cell (col_, row_, '', style_, style_, '', sh_); -- bottom-left
               ELSE
                  Add_Border_To_Cell (col_, row_, '', '', style_, '', sh_); -- left-only
               END IF;
            ELSIF col_ = col_end_ THEN -- last column
               IF row_ = row_start_ THEN
                  Add_Border_To_Cell (col_, row_, style_, '', '', style_, sh_); -- top-right
               ELSIF row_ = row_end_ THEN
                  Add_Border_To_Cell (col_, row_, '', style_, '', style_, sh_); -- bottom-right
               ELSE
                  Add_Border_To_Cell (col_, row_, '', '', '', style_, sh_); -- right-only
               END IF;
            ELSE -- middle columns
               IF row_ = row_start_ THEN
                  Add_Border_To_Cell (col_, row_, style_, '', '', '', sh_); -- top-only
               ELSIF row_ = row_end_ THEN
                  Add_Border_To_Cell (col_, row_, '', style_, '', '', sh_); -- bottom-only
               END IF;
            END IF;

         END LOOP;
      END LOOP;

   END IF;

END Add_Border_To_Range;

PROCEDURE Add_Border_To_Range (
   range_   IN tp_cell_range,
   style_   IN VARCHAR2    := 'medium',
   sheet_   IN PLS_INTEGER := null )
IS
   width_  PLS_INTEGER := range_.tl.c - range_.br.c + 1;
   height_ PLS_INTEGER := range_.tl.r - range_.br.r + 1;
BEGIN
   Add_Border_To_Range (
      range_.tl.c, range_.tl.r, width_, height_, style_, sheet_
   );
END Add_Border_To_Range;

FUNCTION Get_Alignment (
   vertical_   VARCHAR2 := null,
   horizontal_ VARCHAR2 := null,
   wrapText_   BOOLEAN  := null ) RETURN tp_alignment
IS
   rv_ tp_alignment;
BEGIN
   rv_.vertical := vertical_;
   rv_.horizontal := horizontal_;
   rv_.wrapText := wrapText_;
   RETURN rv_;
END Get_Alignment;

FUNCTION Get_Or_Create_XfId (
   Xf_ IN tp_Xf_fmt ) RETURN PLS_INTEGER
IS
   xfId_     PLS_INTEGER;
   Xfi_      tp_Xf_fmt   := Xf_;
   xf_count_ PLS_INTEGER := wb_.cellXfs.count;
   wt_tf_    VARCHAR2(1) := CASE WHEN Xf_.alignment.wrapText THEN 't' ELSE 'f' END;
   md5_hash_ RAW(128)    := Dbms_Crypto.Hash (
      Utl_i18n.String_To_Raw (
         to_char(Xf_.numFmtId) || '^' || to_char(Xf_.fontId) || '^' || to_char(Xf_.fillId) ||
         '^' || to_char(Xf_.borderId) || '^' || nvl (Xf_.alignment.vertical,'x') || '^' ||
         nvl (Xf_.alignment.horizontal,'x') || '^' || wt_tf_,
         'AL32UTF8'
      ),
      dbms_crypto.hash_md5
   );
BEGIN
   FOR i_ IN 1 .. xf_count_ LOOP
      IF wb_.cellXfs(i_).md5 = md5_hash_ THEN
         XfId_ := i_;
         exit;
      END IF;
   END LOOP;
   IF XfId_ IS null THEN -- we didn't find a matching style, so create a new one
      xfId_    := xf_count_ + 1;
      Xfi_.md5 := md5_hash_;
      wb_.cellXfs(xfId_) := Xfi_;
   END IF;
   RETURN xfId_;
END Get_Or_Create_XfId;

FUNCTION Get_XfId (
   sheet_     IN PLS_INTEGER,
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null ) RETURN PLS_INTEGER
IS
   Xf_     tp_Xf_fmt;
   col_Xf_ tp_Xf_fmt;
   row_Xf_ tp_Xf_fmt;
BEGIN

   IF not g_useXf_ THEN
      RETURN null;
   END IF;

   IF wb_.sheets(sheet_).col_fmts.exists(col_) THEN
      col_Xf_ := wb_.sheets(sheet_).col_fmts(col_);
   END IF;
   IF wb_.sheets(sheet_).row_fmts.exists(row_) THEN
      row_Xf_ := wb_.sheets(sheet_).row_fmts(row_);
   END IF;
   Xf_.numFmtId  := coalesce (numFmtId_, col_Xf_.numFmtId, row_Xf_.numFmtId, wb_.sheets(sheet_).fontId, wb_.fontId, 0); -- is this correct with the fontId?
   Xf_.fontId    := coalesce (fontId_, col_Xf_.fontId, row_Xf_.fontId, 0);
   Xf_.fillId    := coalesce (fillId_, col_Xf_.fillId, row_Xf_.fillId, 0);
   Xf_.borderId  := coalesce (borderId_, col_Xf_.borderId, row_Xf_.borderId, 0);
   Xf_.alignment := Get_Alignment (
      coalesce (alignment_.vertical, col_Xf_.alignment.vertical, row_Xf_.alignment.vertical),
      coalesce (alignment_.horizontal, col_Xf_.alignment.horizontal, row_Xf_.alignment.horizontal),
      coalesce (alignment_.wrapText, col_Xf_.alignment.wrapText, row_Xf_.alignment.wrapText)
   );

   IF Xf_.numFmtId + Xf_.fontId + Xf_.fillId + Xf_.borderId = 0
      AND Xf_.alignment.vertical IS null AND Xf_.alignment.horizontal IS null
      AND not nvl(Xf_.alignment.wrapText, false)
   THEN
      RETURN null;
   END IF;

   IF Xf_.numFmtId > 0 THEN
      Set_Col_Width_By_Format (sheet_, col_, Get_Num_Format_Mask(Xf_.numFmtId));
   END IF;

   RETURN Get_Or_Create_XfId (Xf_);

END Get_XfId;

FUNCTION Get_Cell_XfId (
   sheet_ IN PLS_INTEGER,
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER ) RETURN PLS_INTEGER
IS
   style_ PLS_INTEGER;
BEGIN
   IF wb_.sheets(sheet_).rows.exists(row_) AND
      wb_.sheets(sheet_).rows(row_).exists(col_)
   THEN
      style_ := wb_.sheets(sheet_).rows(row_)(col_).style;
   ELSE
      -- We need to create the cell in the PlSql model so that later functions
      -- can manipulate it
      CellB (col_, row_, sheet_ => sheet_);
   END IF;
   RETURN style_;
END Get_Cell_XfId;

-----
-- Get_Cell_Xf()
--   If the cell has an XfId, then we return that Xf without reverting back to
--   rows and columns
--
FUNCTION Get_Cell_Xf (
   sheet_ IN PLS_INTEGER,
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER ) RETURN tp_Xf_fmt
IS
   xfId_ PLS_INTEGER := Get_Cell_XfId (sheet_, col_, row_);
BEGIN
   IF xfId_ IS null THEN
      RETURN null;
   ELSE
      RETURN wb_.cellXfs (xfId_);
   END IF;
END Get_Cell_Xf;

-----
-- Get_Cell_Xff()
--   If the cell doesn't have its own style, then the Xff function goes deeper
--   into the sheet, looking at the column and row styles to see if those also
--   contain values
--
FUNCTION Get_Cell_Xff (
   sheet_ IN PLS_INTEGER,
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER ) RETURN tp_Xf_fmt
IS
   cell_XfId_ PLS_INTEGER := Get_Cell_XfId (sheet_, col_, row_);
   col_Xf_    tp_Xf_fmt;
   row_Xf_    tp_Xf_fmt;
   Xf_        tp_Xf_fmt;
BEGIN

   IF cell_XfId_ IS NOT null THEN
      RETURN wb_.cellXfs (cell_xfId_);
   ELSE

      IF wb_.sheets(sheet_).col_fmts.exists(col_) THEN
         col_Xf_ := wb_.sheets(sheet_).col_fmts(col_);
      END IF;
      IF wb_.sheets(sheet_).row_fmts.exists(row_) THEN
         row_Xf_ := wb_.sheets(sheet_).row_fmts(row_);
      END IF;

      Xf_.numFmtId  := coalesce (col_Xf_.numFmtId, row_Xf_.numFmtId, wb_.sheets(sheet_).fontId, wb_.fontId);  -- is this correct with the fontId??
      Xf_.fontId    := coalesce (col_Xf_.fontId, row_Xf_.fontId, 0);
      Xf_.fillId    := coalesce (col_Xf_.fillId, row_Xf_.fillId, 0);
      Xf_.borderId  := coalesce (col_Xf_.borderId, row_Xf_.borderId, 0);
      Xf_.alignment := Get_Alignment (
         coalesce (col_Xf_.alignment.vertical, row_Xf_.alignment.vertical),
         coalesce (col_Xf_.alignment.horizontal, row_Xf_.alignment.horizontal),
         coalesce (col_Xf_.alignment.wrapText, row_Xf_.alignment.wrapText)
      );
      RETURN Xf_;
   END IF;
END Get_Cell_Xff;


---------------------------------------
---------------------------------------
--
-- Fill Cells with data
--   This group of functions has been through several iterations.  It would be
--   nice to have only 1 `Cell()` function that's overloaded with string, date
--   and number values, but in practice the compiler cannot really distinguish
--   between them effectively.  Hence it's normally better to use the explicit
--   version for each type.
--   We keep the cell's data in a type called `ora_value`; this is useful when
--   the calling program needs to query the data later, or if we want to apply
--   conditional formatting based on that data.
--
--

PROCEDURE Cell ( -- num version
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
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).rows(row_)(col_).datatype  := CELL_DT_NUMBER_;
   wb_.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => '', num_val => value_, dt_val => null
   );
   wb_.sheets(sh_).rows(row_)(col_).value     := value_;
   wb_.sheets(sh_).rows(row_)(col_).style     := get_XfId (
      sh_, col_, row_, numFmtId_, fontId_, fillId_, borderId_, alignment_
   );
END Cell;

PROCEDURE Cell ( -- num version overload
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_num_ IN NUMBER      := null,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null )
IS
   fm_ix_ PLS_INTEGER := wb_.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   Cell (
      col_, row_, value_num_,
      CASE WHEN numFmtId_  IS NOT null THEN numFmt_(numFmtId_) END,
      CASE WHEN fontId_    IS NOT null THEN fonts_(fontId_) END,
      CASE WHEN fillId_    IS NOT null THEN fills_(fillId_) END,
      CASE WHEN borderId_  IS NOT null THEN bdrs_(borderId_) END,
      CASE WHEN alignment_ IS NOT null THEN align_(alignment_) END,
      sheet_
   );
   IF formula_ IS NOT null THEN
      wb_.formulas(fm_ix_) := formula_;
      wb_.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
   END IF;
END Cell;

PROCEDURE CellN ( -- num version explicit
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_num_ IN NUMBER      := null,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null )
IS BEGIN
   Cell (
      col_ => col_, row_ => row_, value_num_ => value_num_, formula_ => formula_,
      numFmtId_ => numFmtId_, fontId_ => fontId_, fillId_ => fillId_,
      borderId_ => borderId_, alignment_ => alignment_, sheet_ => sheet_
   );
END CellN;

FUNCTION Add_String (
   string_ IN VARCHAR2 ) RETURN PLS_INTEGER
IS
   ix_ PLS_INTEGER;
BEGIN
   IF wb_.strings.exists(string_) THEN
      ix_ := wb_.strings(string_);
   ELSE
      ix_ := wb_.strings.count;
      wb_.str_ind(ix_) := string_;
      wb_.strings(string_) := ix_;
   END IF;
   wb_.str_cnt := wb_.str_cnt + 1;
   RETURN ix_;
END Add_String;

FUNCTION Get_Shared_String_Ix (
   string_ IN VARCHAR2 ) RETURN PLS_INTEGER
IS BEGIN
   RETURN CASE
      WHEN not wb_.strings.exists(string_) THEN -1
      ELSE wb_.strings(string_)
   END;
END Get_Shared_String_Ix;

PROCEDURE Cell ( -- string version
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN VARCHAR2,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null )
IS
   sh_    PLS_INTEGER  := nvl(sheet_, wb_.sheets.count);
   align_ tp_alignment := alignment_;
BEGIN
   wb_.sheets(sh_).rows(row_)(col_).datatype  := CELL_DT_STRING_;
   wb_.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => value_, num_val => null, dt_val => null
   );
   wb_.sheets(sh_).rows(row_)(col_).value     := Add_String(value_);
   IF align_.wrapText IS null AND instr(value_, chr(13)) > 0 THEN
      align_.wrapText := true;
   END IF;
   wb_.sheets(sh_).rows(row_)(col_).style := get_XfId (
      sh_, col_, row_, numFmtId_, fontId_, fillId_, borderId_, align_
   );
END Cell;

PROCEDURE Cell ( -- string version overload
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_str_ IN VARCHAR2    := '',
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null ) 
IS
   fm_ix_ PLS_INTEGER := wb_.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   Cell (
      col_, row_, value_str_,
      CASE WHEN numFmtId_  IS NOT null THEN numFmt_(numFmtId_) END,
      CASE WHEN fontId_    IS NOT null THEN fonts_(fontId_) END,
      CASE WHEN fillId_    IS NOT null THEN fills_(fillId_) END,
      CASE WHEN borderId_  IS NOT null THEN bdrs_(borderId_) END,
      CASE WHEN alignment_ IS NOT null THEN align_(alignment_) END,
      sh_
   );
   IF formula_ IS NOT null THEN
      wb_.formulas(fm_ix_) := formula_;
      wb_.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
   END IF;
END Cell;

PROCEDURE CellS ( -- string version explicit
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_str_ IN VARCHAR2    := '',
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null )
IS BEGIN
   Cell (
      col_ => col_, row_ => row_, value_str_ => value_str_, formula_ => formula_,
      numFmtId_ => numFmtId_, fontId_ => fontId_, fillId_ => fillId_,
      borderId_ => borderId_, alignment_ => alignment_, sheet_ => sheet_
   );
END CellS;

FUNCTION Date_To_Xl_Nr (
   date_ IN DATE ) RETURN NUMBER
IS BEGIN
   RETURN date_ - EXCEL_BASE_DATE_;
END Date_To_Xl_Nr;

PROCEDURE Cell (  -- date version
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN DATE,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null )
IS
   num_fmt_id_ PLS_INTEGER := numFmtId_;
   sh_         PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).rows(row_)(col_).datatype  := CELL_DT_DATE_;
   wb_.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => '', num_val => null, dt_val => value_
   );
   wb_.sheets(sh_).rows(row_)(col_).value := Date_To_Xl_Nr(value_);
   IF num_fmt_id_ IS null
      AND not (    wb_.sheets(sh_).col_fmts.exists(col_)
               AND wb_.sheets(sh_).col_fmts(col_).numFmtId IS not null )
      AND not (    wb_.sheets(sh_).row_fmts.exists(row_)
               AND wb_.sheets(sh_).row_fmts(row_).numFmtId IS not null )
   THEN
      num_fmt_id_ := get_numFmt('dd/mm/yyyy');
   END IF;
   wb_.sheets(sh_).rows(row_)(col_).style := get_XfId (
      sh_, col_, row_, num_fmt_id_, fontId_, fillId_, borderId_, alignment_
   );
END Cell;

PROCEDURE Cell ( -- date version overload
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_dt_  IN DATE,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null )
IS
   fm_ix_ PLS_INTEGER := wb_.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   Cell (
      col_, row_, value_dt_,
      CASE WHEN numFmtId_  IS NOT null THEN numFmt_(numFmtId_) END,
      CASE WHEN fontId_    IS NOT null THEN fonts_(fontId_) END,
      CASE WHEN fillId_    IS NOT null THEN fills_(fillId_) END,
      CASE WHEN borderId_  IS NOT null THEN bdrs_(borderId_) END,
      CASE WHEN alignment_ IS NOT null THEN align_(alignment_) END,
      sheet_
   );
   IF formula_ IS NOT null THEN
      wb_.formulas(fm_ix_) := formula_;
      wb_.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
   END IF;
END Cell;

PROCEDURE CellD ( -- date version explicit
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_dt_  IN DATE,
   formula_   IN VARCHAR2    := '',
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null )
IS BEGIN
   Cell (
      col_ => col_, row_ => row_, value_dt_ => value_dt_, formula_ => formula_,
      numFmtId_ => numFmtId_, fontId_ => fontId_, fillId_ => fillId_,
      borderId_ => borderId_, alignment_ => alignment_, sheet_ => sheet_
   );
END CellD;

-- Sometimes it's useful to be able to add an empty cell with formatting
PROCEDURE CellB ( -- empty (b for blank)
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   numFmtId_  IN VARCHAR2    := null,
   fontId_    IN VARCHAR2    := null,
   fillId_    IN VARCHAR2    := null,
   borderId_  IN VARCHAR2    := null,
   alignment_ IN VARCHAR2    := null,
   sheet_     IN PLS_INTEGER := null )
IS BEGIN
   Cell (
      col_, row_, value_str_ => '',
      numFmtId_ => numFmtId_, fontId_ => fontId_, fillId_ => fillId_,
      borderId_ => borderId_, alignment_ => alignment_, sheet_ => sheet_
   );
END CellB;

PROCEDURE Query_Date_Cell (
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER,
   value_ IN DATE,
   sheet_ IN PLS_INTEGER := null,
   XfId_  IN VARCHAR2 )
IS
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   Cell (col_, row_, value_, 0, sheet_ => sheet_);
   wb_.sheets(sh_).rows(row_)(col_).style := XfId_;
END Query_Date_Cell;

--- This function assumes a string value;  perhaps it could be improved...
--- todo
PROCEDURE Condition_Color_Col (
   col_   IN PLS_INTEGER,
   sheet_ IN PLS_INTEGER := null )
IS
   sh_        PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
   first_row_ PLS_INTEGER := wb_.sheets(sh_).rows.first;
   last_row_  PLS_INTEGER := wb_.sheets(sh_).rows.last;
   str_ix_    PLS_INTEGER;
   str_val_   VARCHAR2(50);
   XfId_      PLS_INTEGER;
   num_fmt_   PLS_INTEGER;
   font_id_   PLS_INTEGER;
   border_id_ PLS_INTEGER;
   align_     tp_alignment;

BEGIN

   FOR r_ IN first_row_ .. last_row_ LOOP

      str_ix_  := wb_.sheets(sh_).rows(r_)(col_).value;
      str_val_ := substr (wb_.str_ind(str_ix_), 1, 50);

      IF fills_.exists(str_val_) THEN

         XfId_ := Get_Cell_XfId (sh_, col_, r_);

         IF XfId_ IS null THEN
            wb_.sheets(sh_).rows(r_)(col_).style := get_XfId (
               sh_, col_, r_, fillId_ => fills_(str_val_)
            );
         ELSE
            num_fmt_          := wb_.cellXfs(XfId_).numFmtId;
            font_id_          := wb_.cellXfs(XfId_).fontId;
            border_id_        := wb_.cellXfs(XfId_).borderId;
            align_.vertical   := wb_.cellXfs(XfId_).alignment.vertical;
            align_.horizontal := wb_.cellXfs(XfId_).alignment.horizontal;
            align_.wrapText   := wb_.cellXfs(XfId_).alignment.wrapText;
            wb_.sheets(sh_).rows(r_)(col_).style := get_XfId (
               sh_, col_, r_, num_fmt_, font_id_, fills_(str_val_), border_id_, align_
            );
         END IF;

      END IF;

   END LOOP;

END Condition_Color_Col;

PROCEDURE Hyperlink (
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER,
   url_   IN VARCHAR2,
   value_ IN VARCHAR2    := null,
   sheet_ IN PLS_INTEGER := null )
IS
   ix_  PLS_INTEGER;
   sh_  PLS_INTEGER   := nvl (sheet_, wb_.sheets.count);
   val_ VARCHAR2(200) := nvl (value_, url_);
BEGIN
   wb_.sheets(sh_).rows(row_)(col_).datatype  := CELL_DT_HYPERLINK_;
   wb_.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => val_, num_val => null, dt_val => null
   );
   wb_.sheets(sh_).rows(row_)(col_).value     := Add_String(val_);
   wb_.sheets(sh_).rows(row_)(col_).style     := get_XfId(sh_, col_, row_, '', Get_Font('Calibri', theme_ => 10, underline_ => true));
   ix_ := wb_.sheets(sh_).hyperlinks.count + 1;
   wb_.sheets(sh_).hyperlinks(ix_).cell := Alfan_Cell (col_, row_);
   wb_.sheets(sh_).hyperlinks(ix_).url := url_;
END Hyperlink;


PROCEDURE Comment (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   text_   IN VARCHAR2,
   author_ IN VARCHAR2 := null,
   width_  IN PLS_INTEGER := 150,
   height_ IN PLS_INTEGER := 100,
   sheet_  IN PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
   ix_ PLS_INTEGER := wb_.sheets(sh_).comments.count + 1;
BEGIN
   wb_.sheets(sh_).comments(ix_).row    := row_;
   wb_.sheets(sh_).comments(ix_).column := col_;
   wb_.sheets(sh_).comments(ix_).text   := dbms_xmlgen.convert(text_);
   wb_.sheets(sh_).comments(ix_).author := dbms_xmlgen.convert(author_);
   wb_.sheets(sh_).comments(ix_).width  := width_;
   wb_.sheets(sh_).comments(ix_).height := height_;
END Comment;

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
   sheet_         IN PLS_INTEGER  := null )
IS
   ix_ PLS_INTEGER := wb_.formulas.count;
   sh_ PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   wb_.formulas(ix_) := formula_;
   Cell (col_, row_, default_value_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sh_);
   wb_.sheets(sh_).rows(row_)(col_).formula_idx := ix_;
END Num_Formula;

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
   sheet_         IN PLS_INTEGER  := null )
IS
   ix_ PLS_INTEGER := wb_.formulas.count;
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.formulas(ix_) := formula_;
   Cell (col_, row_, default_value_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sh_);
   wb_.sheets(sh_).rows(row_)(col_).formula_idx := ix_;
END Str_Formula;

PROCEDURE Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN NUMBER      := null,
   numFmtId_      IN VARCHAR2    := null,
   fontId_        IN VARCHAR2    := null,
   fillId_        IN VARCHAR2    := null,
   borderId_      IN VARCHAR2    := null,
   alignment_     IN VARCHAR2    := null,
   sheet_         IN PLS_INTEGER := null )
IS BEGIN
   Cell  (col_, row_, default_value_, formula_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sheet_);
END Formula;

PROCEDURE Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN VARCHAR2    := null,
   numFmtId_      IN VARCHAR2    := null,
   fontId_        IN VARCHAR2    := null,
   fillId_        IN VARCHAR2    := null,
   borderId_      IN VARCHAR2    := null,
   alignment_     IN VARCHAR2    := null,
   sheet_         IN PLS_INTEGER := null )
IS BEGIN
   Cell  (col_, row_, default_value_, formula_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sheet_);
END Formula;

PROCEDURE Formula (
   col_           IN PLS_INTEGER,
   row_           IN PLS_INTEGER,
   formula_       IN VARCHAR2,
   default_value_ IN DATE        := null,
   numFmtId_      IN VARCHAR2    := null,
   fontId_        IN VARCHAR2    := null,
   fillId_        IN VARCHAR2    := null,
   borderId_      IN VARCHAR2    := null,
   alignment_     IN VARCHAR2    := null,
   sheet_         IN PLS_INTEGER := null )
IS BEGIN
   Cell  (col_, row_, default_value_, formula_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sheet_);
END Formula;


PROCEDURE Mergecells (
   tl_col_ IN PLS_INTEGER, -- top left
   tl_row_ IN PLS_INTEGER,
   br_col_ IN PLS_INTEGER, -- bottom right
   br_row_ IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null )
IS
   ix_   PLS_INTEGER;
   sh_ PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   ix_ := wb_.sheets(sh_).mergecells.count + 1;
   wb_.sheets(sh_).mergecells(ix_) := Alfan_Range (tl_col_, tl_row_, br_col_, br_row_);
END Mergecells;

PROCEDURE Add_Validation (
   type_        IN VARCHAR2,
   sqref_       IN VARCHAR2,
   style_       IN VARCHAR2    := 'stop', -- stop, warning, information
   formula1_    IN VARCHAR2    := null,
   formula2_    IN VARCHAR2    := null,
   title_       IN VARCHAR2    := null,
   prompt_      IN VARCHAR     := null,
   show_error_  IN BOOLEAN     := false,
   error_title_ IN VARCHAR2    := null,
   error_txt_   IN VARCHAR2    := null,
   sheet_       IN PLS_INTEGER := null )
IS
   ix_     PLS_INTEGER;
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   ix_ := wb_.sheets(sh_).validations.count + 1;
   wb_.sheets(sh_).validations(ix_).type        := type_;
   wb_.sheets(sh_).validations(ix_).errorstyle  := style_;
   wb_.sheets(sh_).validations(ix_).sqref       := sqref_;
   wb_.sheets(sh_).validations(ix_).formula1    := formula1_;
   wb_.sheets(sh_).validations(ix_).formula2    := formula2_;
   wb_.sheets(sh_).validations(ix_).error_title := error_title_;
   wb_.sheets(sh_).validations(ix_).error_txt   := error_txt_;
   wb_.sheets(sh_).validations(ix_).title       := title_;
   wb_.sheets(sh_).validations(ix_).prompt      := prompt_;
   wb_.sheets(sh_).validations(ix_).showerrormessage := show_error_;
END Add_Validation;

PROCEDURE List_Validation (
   sqref_col_    IN PLS_INTEGER,
   sqref_row_    IN PLS_INTEGER,
   tl_col_       IN PLS_INTEGER, -- top left
   tl_row_       IN PLS_INTEGER,
   br_col_       IN PLS_INTEGER, -- bottom right
   br_row_       IN PLS_INTEGER,
   style_        IN VARCHAR2    := 'stop', -- stop, warning, information
   title_        IN VARCHAR2    := null,
   prompt_       IN VARCHAR     := null,
   show_error_   IN BOOLEAN     := false,
   error_title_  IN VARCHAR2    := null,
   error_txt_    IN VARCHAR2    := null,
   sheet_        IN PLS_INTEGER := null )
IS BEGIN
   Add_Validation (
      type_        => 'list',
      sqref_       => Alfan_Cell (sqref_col_, sqref_row_),
      style_       => lower(style_),
      formula1_    => Alfan_Range (tl_col_, tl_row_, br_col_, br_row_, true, true, true, true),
      title_       => title_,
      prompt_      => prompt_,
      show_error_  => show_error_,
      error_title_ => error_title_,
      error_txt_   => error_txt_,
      sheet_       => sheet_
   );
END List_Validation;

PROCEDURE List_Validation (
   sqref_col_    IN PLS_INTEGER,
   sqref_row_    IN PLS_INTEGER,
   defined_name_ IN VARCHAR2,
   style_        IN VARCHAR2    := 'stop', -- stop, warning, information
   title_        IN VARCHAR2    := null,
   prompt_       IN VARCHAR     := null,
   show_error_   IN BOOLEAN     := false,
   error_title_  IN VARCHAR2    := null,
   error_txt_    IN VARCHAR2    := null,
   sheet_        IN PLS_INTEGER := null )
IS BEGIN
   Add_Validation (
      type_        => 'list',
      sqref_       => Alfan_Cell (sqref_col_, sqref_row_),
      style_       => lower(style_),
      formula1_    => defined_name_,
      title_       => title_,
      prompt_      => prompt_,
      show_error_  => show_error_,
      error_title_ => error_title_,
      error_txt_   => error_txt_,
      sheet_       => sheet_
   );
END List_Validation;

PROCEDURE Add_Image (
   col_         IN PLS_INTEGER,
   row_         IN PLS_INTEGER,
   img_blob_    IN BLOB,
   name_        IN VARCHAR2    := '',
   title_       IN VARCHAR2    := '',
   description_ IN VARCHAR2    := '',
   scale_       IN NUMBER      := null,
   sheet_       IN PLS_INTEGER := null,
   width_       IN PLS_INTEGER := null,
   height_      IN PLS_INTEGER := null )
IS
   sh_         PLS_INTEGER := coalesce (sheet_, wb_.sheets.count);
   img_ix_     PLS_INTEGER;
   hash_       RAW(128) := Dbms_Crypto.Hash (img_blob_, dbms_crypto.hash_md5);
   img_rec_    tp_image;
   drawing_    tp_drawing;
   offset_     NUMBER;
   length_     NUMBER;
   file_chunk_ RAW(14);
   hex_        VARCHAR2(8);
BEGIN

   FOR i_ IN 1 .. wb_.images.count LOOP
      IF wb_.images(i_).img_hash = hash_ THEN
         img_ix_ := i_;
         exit;
      END IF;
   END LOOP;

   IF img_ix_ IS null THEN

      img_ix_ := wb_.images.count + 1;
      dbms_lob.createTemporary (img_rec_.img_blob, true);
      
      dbms_lob.copy (img_rec_.img_blob, img_blob_, dbms_lob.lobmaxsize, 1, 1);
      img_rec_.img_hash := hash_;
      file_chunk_ := dbms_lob.substr (img_blob_, 14, 1);

      --
      -- Different processing for different types of image...
      --
      IF utl_raw.substr (file_chunk_, 1, 8) = hextoraw('89504E470D0A1A0A') THEN -- png
         Dbms_Output.Put_Line ('file is PNG');

         offset_ := 9;
         LOOP
            length_ := to_number (dbms_lob.substr (img_blob_, 4, offset_), 'xxxxxxxx');
            EXIT WHEN length_ IS null OR offset_ > dbms_lob.getlength (img_blob_);
            CASE rawtohex (dbms_lob.substr (img_blob_, 4, offset_ + 4)) -- Chunk type
               WHEN '49484452' /* IHDR */ THEN
                  img_rec_.width  := to_number (dbms_lob.substr(img_blob_,4,offset_+8), 'xxxxxxxx');
                  img_rec_.height := to_number (dbms_lob.substr(img_blob_,4,offset_+12), 'xxxxxxxx');
                  exit;
               WHEN '49454E44' /* IEND */ THEN
                  exit;
            END CASE;
            offset_ := offset_ + 4 + 4 + length_ + 4;  -- Length + Chunk type + Chunk data + CRC
         END LOOP;
         img_rec_.extension := 'png';

      ELSIF utl_raw.substr (file_chunk_, 1, 3) = hextoraw('474946') THEN -- gif
         Dbms_Output.Put_Line ('file is GIF');

         offset_ := 14;
         file_chunk_ := utl_raw.substr (file_chunk_, 11, 1);
         IF utl_raw.bit_and ('80', file_chunk_) = '80' THEN
            length_ := to_number (utl_raw.bit_and('07', file_chunk_), 'XX');
            offset_ := offset_ + 3 * power(2, length_+1);
         END IF;
         LOOP
            CASE rawtohex (dbms_lob.substr (img_blob_, 1, offset_))
               WHEN '21' /* extension */ THEN
                  offset_ := offset_ + 2; -- skip sentinel + label
                  LOOP
                     length_ := to_number(dbms_lob.substr(img_blob_, 1, offset_), 'XX'); -- Block Size
                     EXIT WHEN length_ = 0;
                     offset_ := offset_ + 1 + length_; -- skip Block Size + Data Sub-block
                  END LOOP;
                  offset_ := offset_ + 1; -- skip last Block Size
               WHEN  '2C' /* image */ THEN
                  file_chunk_     := dbms_lob.substr (img_blob_, 4, offset_+5);
                  img_rec_.width  := utl_raw.cast_to_binary_integer (utl_raw.substr(file_chunk_,1,2), utl_raw.little_endian);
                  img_rec_.height := utl_raw.cast_to_binary_integer (utl_raw.substr(file_chunk_,3,2), utl_raw.little_endian);
                  exit;
               ELSE
                  exit;
            END CASE;
         END LOOP;
         img_rec_.extension := 'gif';

      ELSIF utl_raw.substr (file_chunk_,1,2) = hextoraw('FFD8') -- SOI Start of Image
            AND rawtohex (utl_raw.substr(file_chunk_,3,2)) IN ('FFE0', 'FFE1', 'FFEE') -- APP0 jpg; APP1 jpg
      THEN -- jpg
         Dbms_Output.Put_Line ('file is JPG');

         offset_ := 5 + to_number(utl_raw.substr(file_chunk_,5,2), 'xxxx');
         LOOP
            file_chunk_ := dbms_lob.substr (img_blob_, 4, offset_);
            hex_        := substr( rawtohex(file_chunk_),1,4);
            EXIT WHEN hex_ IN ('FFDA', 'FFD9') -- SOS Start of Scan; EOI End Of Image
                   OR substr (hex_, 1, 2) != 'FF';
            IF hex_ IN ('FFD0', 'FFD1', 'FFD2', 'FFD3', 'FFD4', 'FFD5', 'FFD6', 'FFD7', /*RSTn*/ 'FF01' /*TEM*/) THEN
               offset_ := offset_ + 2;
            ELSE
               IF hex_ = 'FFC0' /* SOF0 (Start Of Frame 0) marker*/ THEN
                  hex_ := rawtohex (dbms_lob.substr (img_blob_, 4, offset_+5));
                  img_rec_.width  := to_number (substr(hex_,5), 'xxxx');
                  img_rec_.height := to_number (substr(hex_,1,4), 'xxxx');
                  exit;
               END IF;
               offset_ := offset_ + 2 + to_number (utl_raw.substr(file_chunk_,3,2), 'xxxx');
            END IF;
         END LOOP;
         img_rec_.extension := 'jpeg';

      ELSE -- unknown - use the values passed in
         Dbms_Output.Put_Line ('file is not PNG/GIF/JPG');
         img_rec_.width  := nvl(width_, 0);
         img_rec_.height := nvl(height_, 0);
      END IF;

      wb_.images(img_ix_) := img_rec_;

   END IF;

   drawing_.img_id      := img_ix_;
   drawing_.row         := row_;
   drawing_.col         := col_;
   drawing_.scale       := scale_;
   drawing_.name        := name_;
   drawing_.title       := title_;
   drawing_.description := description_;
   wb_.sheets(sh_).drawings(wb_.sheets(sh_).drawings.count+1) := drawing_;

END Add_Image;

PROCEDURE Load_Image (
   col_         IN PLS_INTEGER,
   row_         IN PLS_INTEGER,
   dir_         IN VARCHAR2,
   filename_    IN VARCHAR2,
   name_        IN VARCHAR2    := '',
   title_       IN VARCHAR2    := '',
   description_ IN VARCHAR2    := '',
   scale_       IN NUMBER      := null,
   sheet_       IN PLS_INTEGER := null,
   width_       IN PLS_INTEGER := null,
   height_      IN PLS_INTEGER := null )
IS
   img_blob_ BLOB  := empty_blob();
   bfile_    BFILE := bFileName (dir_, filename_);
BEGIN
   Dbms_Lob.fileOpen (bfile_);
   Dbms_Lob.createTemporary (img_blob_, true);
   Dbms_Lob.loadFromFile (img_blob_, bfile_, dbms_lob.getLength(bfile_));
   Dbms_Lob.fileClose (bfile_);
   Add_Image (
      col_         => col_,
      row_         => row_,
      img_blob_    => img_blob_,
      name_        => name_,
      title_       => title_,
      description_ => description_,
      scale_       => scale_,
      sheet_       => sheet_,
      width_       => width_,
      height_      => height_
   );
EXCEPTION
   WHEN others THEN
      IF Dbms_Lob.fileIsOpen (bfile_) = 1 THEN
         Dbms_Lob.fileClose (bfile_);
      END IF;
      raise;
END Load_Image;

PROCEDURE Defined_Name (
   name_       VARCHAR2,
   tl_col_     PLS_INTEGER, -- top left
   tl_row_     PLS_INTEGER,
   br_col_     PLS_INTEGER, -- bottom right
   br_row_     PLS_INTEGER,
   fix_tlc_    BOOLEAN     := true,
   fix_tlr_    BOOLEAN     := true,
   fix_brc_    BOOLEAN     := true,
   fix_brr_    BOOLEAN     := true,
   sheet_      PLS_INTEGER := null,
   localsheet_ BOOLEAN     := false )
IS BEGIN
   wb_.defined_names(name_) := tp_cell_range (
      sheet_id     => sheet_,
      tl           => tp_cell_loc (c => tl_col_, r => tl_row_, fixc => fix_tlc_, fixr => fix_tlr_),
      br           => tp_cell_loc (c => br_col_, r => br_row_, fixc => fix_brc_, fixr => fix_brr_),
      defined_name => name_,
      local_sheet  => localsheet_
   );
END Defined_Name;

PROCEDURE Defined_Name (
   range_ IN tp_cell_range )
IS BEGIN
   IF range_.defined_name IS null THEN
      Raise_App_Error ('Defined name cannot be empty!');
   END IF;
   wb_.defined_names(range_.defined_name) := range_;
END Defined_Name;

FUNCTION Range_From_Defined_Name (
   defined_name_ IN VARCHAR2 ) RETURN tp_cell_range
IS BEGIN
   RETURN wb_.defined_names(defined_name_);
END Range_From_Defined_Name;


FUNCTION Create_Pivot_Cache (
   range_    IN tp_cell_range,
   agg_cols_ IN tp_col_agg_fns ) RETURN PLS_INTEGER
IS
   cache_id_ PLS_INTEGER := wb_.pivot_caches.count;
BEGIN
   IF range_.defined_name IS NOT null THEN
      Defined_Name (range_); -- will override existing defined name, so be a little careful
   END IF;
   wb_.pivot_caches(cache_id_) := tp_pivot_cache (
      cache_id      => cache_id_,
      ds_range      => range_,
      flds_to_cache => agg_cols_
   );
   RETURN cache_id_;
END Create_Pivot_Cache;

FUNCTION Get_Agg_Fn_From_Axes (
   pivot_axes_ IN tp_pivot_axes,
   col_ix_     IN PLS_INTEGER ) RETURN VARCHAR2
IS
   rtn_ VARCHAR2(20);
   ix_  PLS_INTEGER;
BEGIN
   rtn_ := CASE WHEN pivot_axes_.col_agg_fns.exists(col_ix_) THEN pivot_axes_.col_agg_fns(col_ix_) END;
   ix_ := pivot_axes_.vrollups.first;
   WHILE ix_ IS NOT null AND rtn_ IS null LOOP
      rtn_ := CASE WHEN pivot_axes_.vrollups(ix_) = col_ix_ THEN 'row' END;
      ix_  := pivot_axes_.vrollups.next(ix_);
   END LOOP;
   ix_ := pivot_axes_.hrollups.first;
   WHILE ix_ IS NOT null AND rtn_ IS null LOOP
      rtn_ := CASE WHEN pivot_axes_.hrollups(ix_) = col_ix_ THEN 'col' END;
      ix_  := pivot_axes_.hrollups.next(ix_);
   END LOOP;
   ix_ := pivot_axes_.filter_cols.first;
   WHILE ix_ IS NOT null AND rtn_ IS null LOOP
      rtn_ := CASE WHEN pivot_axes_.filter_cols(ix_) = col_ix_ THEN 'filter' END;
      ix_  := pivot_axes_.filter_cols.next(ix_);
   END LOOP;
   RETURN rtn_;
END Get_Agg_Fn_From_Axes;

FUNCTION Get_Pivot_Source (
   pivot_id_ IN PLS_INTEGER ) RETURN tp_cell_range
IS
BEGIN
   RETURN wb_.pivot_caches(wb_.pivot_tables(pivot_id_).cache_id).ds_range;
END Get_Pivot_Source;

FUNCTION Add_Pivot_Cache (
   src_data_range_ IN OUT NOCOPY tp_cell_range,
   pivot_axes_     IN tp_pivot_axes ) RETURN PLS_INTEGER
IS
   cols_to_cache_ tp_col_agg_fns;
BEGIN
   Add_Col_Headings_To_Range (src_data_range_); -- easier to access column names later
   FOR c_ IN src_data_range_.col_names.first .. src_data_range_.col_names.last LOOP
      cols_to_cache_(c_) := Get_Agg_Fn_From_Axes (pivot_axes_, c_);
   END LOOP;
   RETURN Create_Pivot_Cache (src_data_range_, cols_to_cache_);
END Add_Pivot_Cache;

PROCEDURE Add_Pivot_Table (
   cache_id_       IN OUT NOCOPY PLS_INTEGER,
   src_data_range_ IN OUT NOCOPY tp_cell_range,
   pivot_axes_     IN tp_pivot_axes,
   location_tl_    IN tp_cell_loc,
   pivot_name_     IN VARCHAR2    := null,
   add_to_sheet_   IN PLS_INTEGER := null,
   new_sheet_name_ IN VARCHAR2    := null )
IS
   pv_id_         PLS_INTEGER := wb_.pivot_tables.count + 1;
   sh_            PLS_INTEGER := CASE
      WHEN add_to_sheet_ IS NOT null THEN add_to_sheet_
      ELSE New_Sheet (nvl (new_sheet_name_, 'Pivot' || pv_id_))
   END;
   sheet_pv_ix_   PLS_INTEGER := wb_.sheets(sh_).pivots_list.count + 1;
BEGIN

   IF cache_id_ IS NOT null AND not wb_.pivot_caches.exists(cache_id_) THEN
      Raise_App_Error ('Cache Id :P1 does not exist in the workbook', cache_id_);
   END IF;

   Add_Col_Headings_To_Range (src_data_range_); -- easier to access column names later

   IF cache_id_ IS null THEN
      cache_id_ := Add_Pivot_Cache (src_data_range_, pivot_axes_);
   -- ELSE???
   --   if the cache already exists, it implies that multiple PTs are using the
   --   same cache.  Should we be updating the `cols_to_cache_` list in this scenario?
   --   THIS FEATURE NOT CURRENTLY SUPPORTED; for now we assume 1:1 cach/pivot
   --   relationship
   END IF;

   wb_.pivot_tables(pv_id_) := tp_pivot_table (
      pivot_table_id => pv_id_,
      pivot_name     => nvl (pivot_name_, 'Pivot' || to_char(pv_id_)),
      cache_id       => cache_id_,
      on_sheet       => sh_,
      location_tl    => location_tl_,
      pivot_axes     => pivot_axes_
   );
   wb_.sheets(sh_).pivots_list(sheet_pv_ix_) := pv_id_;
END Add_Pivot_Table;

PROCEDURE Freeze_Rows (
   nr_rows_ IN PLS_INTEGER := 1,
   sheet_   IN PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).freeze_cols := null;
   wb_.sheets(sh_).freeze_rows := nr_rows_;
END Freeze_Rows;

PROCEDURE Freeze_Cols (
   nr_cols_ IN PLS_INTEGER := 1,
   sheet_   IN PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).freeze_rows := null;
   wb_.sheets(sh_).freeze_cols := nr_cols_;
END Freeze_Cols;

PROCEDURE Freeze_Pane (
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER,
   sheet_ IN PLS_INTEGER := null )
IS
   sh_ PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).freeze_rows := row_;
   wb_.sheets(sh_).freeze_cols := col_;
END Freeze_Pane;

PROCEDURE Set_Autofilter (
   col_start_ IN PLS_INTEGER := null,
   col_end_   IN PLS_INTEGER := null,
   row_start_ IN PLS_INTEGER := null,
   row_end_   IN PLS_INTEGER := null,
   sheet_     IN PLS_INTEGER := null )
IS
   ix_ PLS_INTEGER := 1;
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).autofilters(ix_).column_start := col_start_;
   wb_.sheets(sh_).autofilters(ix_).column_end   := col_end_;
   wb_.sheets(sh_).autofilters(ix_).row_start    := row_start_;
   wb_.sheets(sh_).autofilters(ix_).row_end      := row_end_;
   Defined_Name (
      '_xlnm._FilterDatabase', col_start_, row_start_, col_end_, row_end_,
      false, false, false, false, sh_
   );
END Set_Autofilter;


---------------------------------------
---------------------------------------
--
-- The Excel file's XML creators
--
--
PROCEDURE Finish_Content_Types (
   excel_ IN OUT NOCOPY BLOB )
IS
   s_         PLS_INTEGER;
   doc_       dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_types_  dbms_XmlDom.DomNode;
   attrs_     xml_attrs_arr;
   img_exts_  tp_strings;
   ext_       VARCHAR2(5);
   pt_        PLS_INTEGER;
BEGIN

   -- [Content_Types].xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types', attrs_);
   nd_types_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Types', attrs_);

   IF wb_.images.count > 0 THEN
      FOR img_ IN wb_.images.first .. wb_.images.last LOOP
         ext_ := wb_.images(img_).extension;
         IF ext_ IS NOT null AND not img_exts_.exists(ext_) THEN
            natr ('ContentType', 'image/' || ext_, attrs_);
            attr ('Extension', ext_, attrs_);
            Xml_Node (doc_, nd_types_, 'Default', attrs_);
            img_exts_(ext_) := 1;
         END IF;
      END LOOP;
   END IF;

   natr ('Extension',   'rels', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-package.relationships+xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Default', attrs_);

   natr ('Extension',   'xml', attrs_);
   attr ('ContentType', 'application/xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Default', attrs_);

   natr ('Extension',   'vml', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-officedocument.vmlDrawing', attrs_);
   Xml_Node (doc_, nd_types_, 'Default', attrs_);

   natr ('PartName', '/xl/workbook.xml', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Override', attrs_);

   FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
      natr ('PartName',    rep('/xl/pivotCache/pivotCacheDefinition:P1.xml', pc_), attrs_);
      attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml', attrs_);
      Xml_Node (doc_, nd_types_, 'Override', attrs_);
      natr ('PartName', rep('/xl/pivotCache/pivotCacheRecords:P1.xml', pc_), attrs_);
      attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml', attrs_);
      Xml_Node (doc_, nd_types_, 'Override', attrs_);
   END LOOP;
   FOR pt_ IN 1 .. wb_.pivot_tables.count LOOP
      natr ('PartName', rep('/xl/pivotTables/pivotTable:P1.xml', pt_), attrs_);
      attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml', attrs_);
      Xml_Node (doc_, nd_types_, 'Override', attrs_);
   END LOOP;

   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      natr ('PartName', rep('/xl/worksheets/sheet:P1.xml', to_char(s_)), attrs_);
      attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml', attrs_);
      Xml_Node (doc_, nd_types_, 'Override', attrs_);
      s_ := wb_.sheets.next(s_);
   END LOOP;

   natr ('PartName', '/xl/theme/theme1.xml', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-officedocument.theme+xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Override', attrs_);
   natr ('PartName', '/xl/styles.xml', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Override', attrs_);
   natr ('PartName', '/xl/sharedStrings.xml', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Override', attrs_);

   natr ('PartName', '/docProps/core.xml', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-package.core-properties+xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Override', attrs_);
   natr ('PartName', '/docProps/app.xml', attrs_);
   attr ('ContentType', 'application/vnd.openxmlformats-officedocument.extended-properties+xml', attrs_);
   Xml_Node (doc_, nd_types_, 'Override', attrs_);

   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      IF wb_.sheets(s_).comments.count > 0 THEN
         natr ('PartName', rep('/xl/comments:P1.xml', s_), attrs_);
         attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml', attrs_);
         Xml_Node (doc_, nd_types_, 'Override', attrs_);
      END IF;
      IF wb_.sheets(s_).drawings.count > 0 THEN
         natr ('PartName', rep('/xl/drawings/drawing:P1.xml', s_), attrs_);
         attr ('ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml', attrs_);
         Xml_Node (doc_, nd_types_, 'Override', attrs_);
      END IF;
      s_ := wb_.sheets.next(s_);
   END LOOP;

   Add1Xml (excel_, '[Content_Types].xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Content_Types;

PROCEDURE Finish_Rels (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_rels_  dbms_XmlDom.DomNode;
   attrs_    xml_attrs_arr;
BEGIN

   -- _rels/.rels
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rels_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

   natr ('Id', 'rId1', attrs_);
   attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', attrs_);
   attr ('Target', 'xl/workbook.xml', attrs_);
   Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
   natr ('Id', 'rId2', attrs_);
   attr ('Type', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', attrs_);
   attr ('Target', 'docProps/core.xml', attrs_);
   Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
   natr ('Id', 'rId3', attrs_);
   attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', attrs_);
   attr ('Target', 'docProps/app.xml', attrs_);
   Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);

   Add1Xml (excel_, '_rels/.rels', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Rels;

PROCEDURE Finish_docProps (
   excel_ IN OUT NOCOPY BLOB )
IS
   s_        PLS_INTEGER;
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_cprop_ dbms_XmlDom.DomNode;
   nd_prop_  dbms_XmlDom.DomNode;
   nd_hd_    dbms_XmlDom.DomNode;
   nd_vec_   dbms_XmlDom.DomNode;
   nd_var_   dbms_XmlDom.DomNode;
   nd_top_   dbms_XmlDom.DomNode;
   attrs_    xml_attrs_arr;
BEGIN

   -- docProps/core.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', attrs_);
   attr ('xmlns:dc', 'http://purl.org/dc/elements/1.1/', attrs_);
   attr ('xmlns:dcterms', 'http://purl.org/dc/terms/', attrs_);
   attr ('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', attrs_);
   attr ('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', attrs_);
   nd_cprop_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'coreProperties', 'cp', attrs_);

   Xml_Text_Node (doc_, nd_cprop_, 'creator',        sys_context('userenv','os_user'), 'dc');
   Xml_Text_Node (doc_, nd_cprop_, 'description',    rep('Build by version: :P1', VERSION_), 'dc');
   Xml_Text_Node (doc_, nd_cprop_, 'lastModifiedBy', sys_context('userenv','os_user'), 'cp');

   natr ('xsi:type', 'dcterms:W3CDTF', attrs_);
   Xml_Text_Node (doc_, nd_cprop_, 'created',  to_char(current_timestamp,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM'), 'dcterms', attrs_);
   Xml_Text_Node (doc_, nd_cprop_, 'modified', to_char(current_timestamp,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM'), 'dcterms', attrs_);

   Add1Xml (excel_, 'docProps/core.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);


   -- docProps/app.xml
   doc_ := Dbms_XmlDom.newDomDocument;
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   natr ('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties', attrs_);
   attr ('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes', attrs_);
   nd_prop_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Properties', attrs_);

   Xml_Text_Node (doc_, nd_prop_, 'Application', 'Microsoft Excel');
   Xml_Text_Node (doc_, nd_prop_, 'DocSecurity', '0');
   Xml_Text_Node (doc_, nd_prop_, 'ScaleCrop', 'false');
   nd_hd_  := Xml_Node (doc_, nd_prop_, 'HeadingPairs');

   natr ('size',     '2', attrs_);
   attr ('baseType', 'variant', attrs_);
   nd_vec_ := Xml_Node (doc_, nd_hd_, 'vector', 'vt', attrs_);
   nd_var_ := Xml_Node (doc_, nd_vec_, 'variant', 'vt');
   Xml_Text_Node (doc_, nd_var_, 'lpstr', 'Worksheets', 'vt');
   nd_var_ := Xml_Node (doc_, nd_vec_, 'variant', 'vt');
   Xml_Text_Node (doc_, nd_var_, 'i4', to_char(wb_.sheets.count), 'vt');

   nd_top_ := Xml_Node (doc_, nd_prop_, 'TitlesOfParts');
   natr ('size', wb_.sheets.count, attrs_);
   attr ('baseType', 'lpstr', attrs_);
   nd_vec_ := Xml_Node (doc_, nd_top_, 'vector', 'vt', attrs_);
   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      Xml_Text_Node (doc_, nd_vec_, 'lpstr', wb_.sheets(s_).name, 'vt');
      s_ := wb_.sheets.next(s_);
   END LOOP;
   Xml_Text_Node (doc_, nd_prop_, 'LinksUpToDate', 'false');
   Xml_Text_Node (doc_, nd_prop_, 'SharedDoc', 'false');
   Xml_Text_Node (doc_, nd_prop_, 'HyperlinksChanged', 'false');
   Xml_Text_Node (doc_, nd_prop_, 'AppVersion', '14.0300');

   Add1Xml (excel_, 'docProps/app.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_docProps;

PROCEDURE Finish_Shared_Strings (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_    dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_sst_ dbms_XmlDom.DomNode;
   attrs_  xml_attrs_arr;
BEGIN

   -- xl/sharedStrings.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   attr ('count', to_char(wb_.str_cnt), attrs_);
   attr ('uniqueCount', wb_.strings.count, attrs_);
   nd_sst_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'sst', attrs_);

   natr ('xml:space', 'preserve', attrs_);
   FOR str_ix_ IN 0 .. wb_.str_ind.count - 1 LOOP
      Xml_Text_Node (
         doc_ => doc_, append_to_ => Xml_Node(doc_,nd_sst_,'si'), tag_name_ => 't',
         text_content_ => Dbms_XmlGen.Convert (substr(wb_.str_ind(str_ix_), 1, 32000)),
         attrs_ => attrs_
      );
   END LOOP;

   Add1Xml (excel_, 'xl/sharedStrings.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Shared_Strings;

PROCEDURE Finish_Styles (
   excel_ IN OUT NOCOPY BLOB )
IS
   format_mask_ VARCHAR2(100);
   doc_         dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_stl_      dbms_XmlDom.DomNode;
   nd_numf_     dbms_XmlDom.DomNode;
   nd_fnts_     dbms_XmlDom.DomNode;
   nd_fnt_      dbms_XmlDom.DomNode;
   nd_fills_    dbms_XmlDom.DomNode;
   nd_fill_     dbms_XmlDom.DomNode;
   nd_bdrs_     dbms_XmlDom.DomNode;
   nd_bdr_      dbms_XmlDom.DomNode;
   nd_pf_       dbms_XmlDom.DomNode;
   nd_sxfs_     dbms_XmlDom.DomNode;
   nd_xfs_      dbms_XmlDom.DomNode;
   nd_xf_       dbms_XmlDom.DomNode;
   attrs_       xml_attrs_arr;
BEGIN

   -- xl/styles.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   attr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
   attr ('mc:Ignorable', 'x14ac', attrs_);
   attr ('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', attrs_);
   nd_stl_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'styleSheet', attrs_);

   IF wb_.numFmts.count > 0 THEN
      natr ('count', to_char(wb_.numFmts.count), attrs_);
      nd_numf_ := Xml_Node (doc_, nd_stl_, 'numFmts', attrs_);
      format_mask_ := wb_.numFmts.first;
      WHILE format_mask_ IS NOT null LOOP
         natr ('numFmtId', wb_.numFmts(format_mask_), attrs_);
         attr ('formatCode', format_mask_, attrs_);
         Xml_Node (doc_, nd_numf_, 'numFmt', attrs_);
         format_mask_ := wb_.numFmts.next(format_mask_);
      END LOOP;
   END IF;

   natr ('count', wb_.fonts.count, attrs_);
   attr ('x14ac:knownFonts', '1', attrs_);
   nd_fnts_ := Xml_Node (doc_, nd_stl_, 'fonts', attrs_);
   FOR f_ IN 0 .. wb_.fonts.count-1 LOOP
      nd_fnt_ := Xml_Node (doc_, nd_fnts_, 'font');
      IF wb_.fonts(f_).bold     THEN Xml_Node (doc_, nd_fnt_, 'b'); END IF;
      IF wb_.fonts(f_).italic   THEN Xml_Node (doc_, nd_fnt_, 'i'); END IF;
      IF wb_.fonts(f_).underline THEN Xml_Node (doc_, nd_fnt_, 'u'); END IF;

      natr ('val', to_char(wb_.fonts(f_).fontsize, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'), attrs_);
      Xml_Node (doc_, nd_fnt_, 'sz', attrs_);
      attrs_.delete;
      IF wb_.fonts(f_).rgb IS NOT null THEN
         attr ('rgb', wb_.fonts(f_).rgb, attrs_);
      ELSE
         attr ('theme', wb_.fonts(f_).theme, attrs_);
      END IF;
      Xml_Node (doc_, nd_fnt_, 'color', attrs_);

      natr ('val', wb_.fonts(f_).name, attrs_);
      Xml_Node (doc_, nd_fnt_, 'name', attrs_);
      natr ('val', wb_.fonts(f_).family, attrs_);
      Xml_Node (doc_, nd_fnt_, 'family', attrs_);
      natr ('val', 'none', attrs_);
      Xml_Node (doc_, nd_fnt_, 'scheme', attrs_);
   END LOOP;

   natr ('count', wb_.fills.count, attrs_);
   nd_fills_ := Xml_Node (doc_, nd_stl_, 'fills', attrs_);
   FOR f_ IN 0 .. wb_.fills.count-1 LOOP
      nd_fill_ := Xml_Node (doc_, nd_fills_, 'fill');
      natr ('patternType', wb_.fills(f_).patternType, attrs_);
      nd_pf_ := Xml_Node (doc_, nd_fill_, 'patternFill', attrs_);
      attrs_.delete;
      IF wb_.fills(f_).fgRGB IS NOT null THEN
         attr ('rgb', wb_.fills(f_).fgRGB, attrs_);
         Xml_Node (doc_, nd_pf_, 'fgColor', attrs_);
      END IF;
      IF wb_.fills(f_).bgRGB IS NOT null THEN
         attr ('rgb', wb_.fills(f_).bgRGB, attrs_);
         Xml_Node (doc_, nd_pf_, 'bgColor', attrs_);
      END IF;
   END LOOP;

   natr ('count', wb_.borders.count, attrs_);
   nd_bdrs_ := Xml_Node (doc_, nd_stl_, 'borders', attrs_);
   FOR b_ IN 0 .. wb_.borders.count-1 LOOP
      nd_bdr_ := Xml_Node (doc_, nd_bdrs_, 'border');
      attrs_.delete;
      IF wb_.borders(b_).left   IS null THEN attrs_.delete; ELSE attr('style', wb_.borders(b_).left, attrs_); END IF;
      Xml_Node (doc_, nd_bdr_, 'left', attrs_);
      IF wb_.borders(b_).right  IS null THEN attrs_.delete; ELSE attr('style', wb_.borders(b_).right, attrs_); END IF;
      Xml_Node (doc_, nd_bdr_, 'right', attrs_);
      IF wb_.borders(b_).top    IS null THEN attrs_.delete; ELSE attr('style', wb_.borders(b_).top, attrs_); END IF;
      Xml_Node (doc_, nd_bdr_, 'top', attrs_);
      IF wb_.borders(b_).bottom IS null THEN attrs_.delete; ELSE attr('style', wb_.borders(b_).bottom, attrs_); END IF;
      Xml_Node (doc_, nd_bdr_, 'bottom', attrs_);
   END LOOP;

   natr ('count', '1', attrs_);
   nd_sxfs_ := Xml_Node (doc_, nd_stl_, 'cellStyleXfs', attrs_);
   natr ('numFmtId', '0', attrs_);
   attr ('fontId', '0', attrs_);
   attr ('fillId', '0', attrs_);
   attr ('borderId', '0', attrs_);
   Xml_Node (doc_, nd_sxfs_, 'xf', attrs_);

   natr ('count', wb_.cellXfs.count+1, attrs_);
   nd_xfs_ := Xml_Node (doc_, nd_stl_, 'cellXfs', attrs_);

   natr ('numFmtId', '0', attrs_);
   attr ('fontId', '0', attrs_);
   attr ('fillId', '0', attrs_);
   attr ('borderId', '0', attrs_);
   attr ('xfId', '0', attrs_);
   Xml_Node (doc_, nd_xfs_, 'xf', attrs_);
   FOR x_ IN 1 .. wb_.cellXfs.count LOOP
      attrs_.delete;
      natr ('numFmtId', wb_.cellXfs(x_).numFmtId, attrs_);
      attr ('fontId', wb_.cellXfs(x_).fontId, attrs_);
      attr ('fillId', wb_.cellXfs(x_).fillId, attrs_);
      attr ('borderId', wb_.cellXfs(x_).borderId, attrs_);
      nd_xf_ := Xml_Node (doc_, nd_xfs_, 'xf', attrs_);
      IF wb_.cellXfs(x_).alignment.horizontal IS NOT null OR wb_.cellXfs(x_).alignment.vertical IS NOT null OR wb_.cellXfs(x_).alignment.wrapText IS NOT null THEN
         attrs_.delete;
         IF wb_.cellXfs(x_).alignment.horizontal IS NOT null THEN attr('horizontal', wb_.cellXfs(x_).alignment.horizontal, attrs_); END IF;
         IF wb_.cellXfs(x_).alignment.vertical    IS NOT null THEN attr('vertical', wb_.cellXfs(x_).alignment.vertical, attrs_); END IF;
         IF wb_.cellXfs(x_).alignment.wrapText THEN attr('wrapText', 'true', attrs_); END IF;
         Xml_Node (doc_, nd_xf_, 'alignment', attrs_);
      END IF;
   END LOOP;

   natr ('count', '1', attrs_);
   nd_xfs_ := Xml_Node (doc_, nd_stl_, 'cellStyles', attrs_);

   natr ('name', 'Normal', attrs_);
   attr ('xfId', '0', attrs_);
   attr ('builtinId', '0', attrs_);
   Xml_Node (doc_, nd_xfs_, 'cellStyle', attrs_);

   natr ('count', '0', attrs_);
   Xml_Node (doc_, nd_stl_, 'dxfs', attrs_);
   natr ('defaultTableStyle', 'TableStyleMedium2', attrs_);
   attr ('defaultPivotStyle', 'PivotStyleLight16', attrs_);
   Xml_Node (doc_, nd_stl_, 'tableStyles', attrs_);

   nd_xfs_ := Xml_Node (doc_, nd_stl_, 'extLst');
   natr ('uri', '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}', attrs_);
   attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
   nd_xf_ := Xml_Node (doc_, nd_xfs_, 'ext', attrs_);
   natr ('defaultSlicerStyle', 'SlicerStyleLight1', attrs_);
   Xml_Node (doc_, nd_xf_, 'slicerStyles', 'x14', attrs_);

   Add1Xml (excel_, 'xl/styles.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Styles;


PROCEDURE Finish_Theme (
   excel_ IN OUT NOCOPY BLOB )
IS BEGIN
   -- xl/theme/theme1.xml
   Add1Xml (excel_, 'xl/theme/theme1.xml',
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/>
        <a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/>
        <a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/>
        <a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/>
        <a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/>
        <a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/>
        <a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/>
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
</a:theme>');
END Finish_Theme;


PROCEDURE Finish_Workbook (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_wb_    dbms_XmlDom.DomNode;
   nd_bks_   dbms_XmlDom.DomNode;
   nd_shs_   dbms_XmlDom.DomNode;
   nd_dnm_   dbms_XmlDom.DomNode;
   nd_pvs_   dbms_XmlDom.DomNode;
   nd_extl_  dbms_XmlDom.DomNode;
   nd_ext_   dbms_XmlDom.DomNode;
   nd_cf_    dbms_XmlDom.DomNode;
   attrs_    xml_attrs_arr;
   s_        PLS_INTEGER;
   dn_       VARCHAR2(100);
   rel_      PLS_INTEGER := 4; -- see hard-coded rels in Finish_Workbook_Rels()
BEGIN

   -- xl/workbook.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
   nd_wb_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'workbook', attrs_);

   natr ('appName', 'xl', attrs_);
   attr ('lastEdited', '5', attrs_);
   attr ('lowestEdited', '5', attrs_);
   attr ('rupBuild', '9302', attrs_);
   Xml_Node (doc_, nd_wb_, 'fileVersion', attrs_);
   attrs_.delete;
   IF wb_.pivot_tables.count > 0 THEN
      attr ('hidePivotFieldList', '1', attrs_);
   END IF;
   attr ('defaultThemeVersion', '166925', attrs_);
   attr ('date1904', 'false', attrs_);
   Xml_Node (doc_, nd_wb_, 'workbookPr', attrs_);

   nd_bks_ := Xml_Node (doc_, nd_wb_, 'bookViews');
   natr ('xWindow',  '120', attrs_);
   attr ('yWindow', '45', attrs_);
   attr ('windowWidth', '19155', attrs_);
   attr ('windowHeight', '4935', attrs_);
   Xml_Node (doc_, nd_bks_, 'workbookView', attrs_);

   nd_shs_ := Xml_Node (doc_, nd_wb_, 'sheets');
   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      natr ('name', wb_.sheets(s_).name, attrs_);
      attr ('sheetId', to_char(s_), attrs_);
      attr ('r:id', rep ('rId:P1', to_char(rel_)), attrs_);
      Xml_Node (doc_, nd_shs_, 'sheet', attrs_);
      wb_.sheets(s_).wb_rel := rel_;
      rel_ := rel_ + 1;
      s_   := wb_.sheets.next(s_);
   END LOOP;

   IF wb_.defined_names.count > 0 THEN
      nd_dnm_ := Xml_Node (doc_, nd_wb_, 'definedNames');
      dn_ := wb_.defined_names.first;
      WHILE dn_ IS NOT null LOOP
         natr ('name', dn_, attrs_);
         IF wb_.defined_names(dn_).local_sheet THEN
            IF wb_.defined_names(dn_).sheet_id IS null THEN
               Raise_App_Error ('Sheet Id must be defined for local-sheet function to be viable!');
            END IF;
            attr ('localSheetId', to_char(wb_.defined_names(dn_).sheet_id), attrs_);
         END IF;
         Xml_Text_Node (doc_, nd_dnm_, 'definedName', Alfan_Sheet_Range(wb_.defined_names(dn_)), attrs_);
         dn_ := wb_.defined_names.next(dn_);
      END LOOP;
   END IF;

   natr ('calcId', '144525', attrs_);
   Xml_Node (doc_, nd_wb_, 'calcPr', attrs_);

   IF wb_.pivot_caches.count > 0 THEN
      nd_pvs_ :=  Xml_Node (doc_, nd_wb_, 'pivotCaches');
      FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
         natr ('cacheId', to_char(wb_.pivot_caches(pc_).cache_id), attrs_);
         attr ('r:id', 'rId' || to_char(rel_), attrs_);
         Xml_Node (doc_, nd_pvs_, 'pivotCache', attrs_);
         wb_.pivot_caches(pc_).wb_rel := rel_;
         rel_                         := rel_ + 1;
      END LOOP;

      nd_extl_ := Xml_Node (doc_, nd_wb_, 'extLst');

      natr ('uri', Get_Guid, attrs_);
      attr ('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main', attrs_);
      nd_ext_ := Xml_Node (doc_, nd_extl_, 'ext', attrs_);

      natr ('chartTrackingRefBase', '1', attrs_);
      Xml_Node (doc_, nd_ext_, 'workbookPr', 'x15', attrs_);

      natr ('uri', Get_Guid, attrs_);
      attr ('xmlns:xcalcf', 'http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures', attrs_);
      nd_ext_ := Xml_Node (doc_, nd_extl_, 'ext', attrs_);

      nd_cf_ := Xml_Node (doc_, nd_ext_, 'calcFeatures', 'xcalcf');

      natr ('name', 'microsoft.com:RD', attrs_);
      Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      natr ('name', 'microsoft.com:Single', attrs_);
      Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      natr ('name', 'microsoft.com:FV', attrs_);
      Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      natr ('name', 'microsoft.com:CNMTM', attrs_);
      Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      natr ('name', 'microsoft.com:LET_WF', attrs_);
      Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);

   END IF;

   Add1Xml (excel_, 'xl/workbook.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Workbook;

PROCEDURE Finish_Workbook_Rels (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_    dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_rls_ dbms_XmlDom.DomNode;
   attrs_  xml_attrs_arr;
   s_      PLS_INTEGER;
BEGIN

   -- xl/_rels/workbook.xml.rels
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rls_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

   natr ('Id', 'rId1', attrs_);
   attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', attrs_);
   attr ('Target', 'sharedStrings.xml', attrs_);
   Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);

   natr ('Id', 'rId2', attrs_);
   attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', attrs_);
   attr ('Target', 'styles.xml', attrs_);
   Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);

   natr ('Id', 'rId3', attrs_);
   attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', attrs_);
   attr ('Target', 'theme/theme1.xml', attrs_);
   Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);

   FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
      natr ('Id', 'rId' || to_char (wb_.pivot_caches(pc_).wb_rel), attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition', attrs_);
      attr ('Target', rep ('pivotCache/pivotCacheDefinition:P1.xml', pc_), attrs_);
      Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);
   END LOOP;

   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      natr ('Id', 'rId' || to_char(wb_.sheets(s_).wb_rel), attrs_);
      attr ('Type',  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', attrs_);
      attr ('Target', rep ('worksheets/sheet:P1.xml', to_char(s_)), attrs_);
      Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);
      s_ := wb_.sheets.next(s_);
   END LOOP;

   Add1Xml (excel_, 'xl/_rels/workbook.xml.rels', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Workbook_Rels;

PROCEDURE Build_Pivot_Caches
IS
   rollup_fn_      VARCHAR2(20);
   col_name_       VARCHAR2(32000);
   range_vals_ord_ tp_data_ix_ord;
   cache_field_    tp_cache_field;
   min_val_        NUMBER;
   max_val_        NUMBER;
   uq_             tp_unique_data;
BEGIN
   FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
      FOR c_ IN 1 .. Range_Width (wb_.pivot_caches(pc_).ds_range) LOOP

         col_name_  := Range_Col_Head_Name (wb_.pivot_caches(pc_).ds_range, c_);
         rollup_fn_ := wb_.pivot_caches(pc_).flds_to_cache(c_);

         IF rollup_fn_ IN ('row','col','filter') THEN -- filter needs to be checked!!
            range_vals_ord_ := Range_Unique_Data_Ord (wb_.pivot_caches(pc_).ds_range, c_);
            uq_ := Ord_Data_To_Unique (range_vals_ord_);
            cache_field_ := tp_cache_field (
               field_name   => col_name_,
               rollup_fn    => rollup_fn_,
               format_id    => Range_Col_NumFmtId (wb_.pivot_caches(pc_).ds_range, c_),
               shared_items => uq_,
               si_order     => range_vals_ord_,
               min_value    => null,
               max_value    => null
            );
         ELSIF rollup_fn_ IN ('sum') THEN
            Range_Col_Min_Max_Values (
               wb_.pivot_caches(pc_).ds_range, c_, min_val_, max_val_
            );
            cache_field_ := tp_cache_field (
               field_name   => col_name_,
               rollup_fn    => rollup_fn_,
               format_id    => Range_Col_NumFmtId (wb_.pivot_caches(pc_).ds_range, c_),
               shared_items => tp_unique_data(),
               si_order     => tp_data_ix_ord(),
               min_value    => min_val_,
               max_value    => max_val_
            );
         ELSE
            cache_field_ := tp_cache_field (
               field_name   => col_name_,
               rollup_fn    => '',
               format_id    => 0,
               shared_items => tp_unique_data(),
               si_order     => tp_data_ix_ord(),
               min_value    => null,
               max_value    => null
            );
         END IF;
         wb_.pivot_caches(pc_).cached_fields(col_name_) := cache_field_;
         wb_.pivot_caches(pc_).cf_order(c_)             := col_name_;
      END LOOP;
   END LOOP;
END Build_Pivot_Caches;

PROCEDURE Finish_Pivot_Caches (
   excel_ IN OUT NOCOPY BLOB )
IS
   sh_      PLS_INTEGER;
   doc_     dbms_XmlDom.DomDocument;
   attrs_   xml_attrs_arr;
   nd_pcd_  dbms_XmlDom.DomNode;
   nd_cs_   dbms_XmlDom.DomNode;
   nd_cfs_  dbms_XmlDom.DomNode;
   nd_cf_   dbms_XmlDom.DomNode;
   nd_rels_ dbms_XmlDom.DomNode;
   nd_si_   dbms_XmlDom.DomNode;
   nd_el_   dbms_XmlDom.DomNode;
   nd_ex_   dbms_XmlDom.DomNode;
   nd_row_  dbms_XmlDom.DomNode;
   tag_     VARCHAR2(1);
   fld_     VARCHAR2(2000);
   cache_   tp_pivot_cache;
   cfld_    tp_cache_field;
   xl_col_  PLS_INTEGER;
BEGIN

   FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP

      cache_ := wb_.pivot_caches(pc_);
      sh_ := cache_.ds_range.sheet_id;
      IF sh_ IS null THEN
         Raise_App_Error ('A data-range must have a sheet defined inside a cache definition, in Finish_Pivot_Caches()');
      END IF;

      -- xl/pivotCache/pivotCacheDefinition:P1.xml
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
      attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
      attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
      attr ('mc:Ignorable', 'xr', attrs_);
      attr ('r:id', 'rId1', attrs_); -- points to pivot-cache-record; there's only ever 1 per pCD
      attr ('refreshedBy', user, attrs_);
      attr ('refreshedDate', Date_To_Xl_Nr(sysdate), attrs_);
      attr ('createdVersion', '7', attrs_); -- Version of Excel in which this pivot was created!
      attr ('refreshedVersion', '7', attrs_);
      attr ('minRefreshableVersion', '3', attrs_); -- Minimum version of Excel which is compatible (apparently)
      attr ('recordCount', to_char(Range_Height(cache_.ds_range)), attrs_);
      attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
      attr ('xr:uid', Get_Guid, attrs_); --'{C898DCD4-A18D-452F-B655-4FAEB857F78F}';
      nd_pcd_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'pivotCacheDefinition', attrs_);

      natr ('type', 'worksheet', attrs_);
      nd_cs_ := Xml_Node (doc_, nd_pcd_, 'cacheSource', attrs_);

      attrs_.delete;
      IF cache_.ds_range.defined_name IS NOT null THEN
         attr ('name', cache_.ds_range.defined_name, attrs_);
      ELSE
         attr ('ref', Alfan_Range (cache_.ds_range), attrs_);
         attr ('sheet', Sheet_Name (cache_.ds_range), attrs_);
      END IF;
      Xml_Node (doc_, nd_cs_, 'worksheetSource', attrs_);

      attrs_.delete;
      attr ('count', to_char(cache_.cf_order.count), attrs_);
      nd_cfs_ := Xml_Node (doc_, nd_pcd_, 'cacheFields', attrs_);

      FOR c_ IN cache_.cf_order.first .. cache_.cf_order.last LOOP

         fld_  := cache_.cf_order(c_);
         cfld_ := cache_.cached_fields(fld_);

         natr ('name', fld_, attrs_);
         attr ('numFmtId', cache_.cached_fields(fld_).format_id, attrs_);
         nd_cf_ := Xml_Node (doc_, nd_cfs_, 'cacheField', attrs_);

         attrs_.delete;
         IF cfld_.rollup_fn IN ('row','column','fileter') THEN
            attr ('count', to_char(cache_.cached_fields(fld_).shared_items.count), attrs_);
         ELSIF cfld_.rollup_fn = 'sum' THEN
            attr ('containsSemiMixedTypes', '0', attrs_);
            attr ('containsString', '0', attrs_);
            attr ('containsNumber', '1', attrs_);
            attr ('minValue', to_char(cache_.cached_fields(fld_).min_value), attrs_);
            attr ('maxValue', to_char(cache_.cached_fields(fld_).max_value), attrs_);
         END IF;
         nd_si_ := Xml_Node (doc_, nd_cf_, 'sharedItems', attrs_);

         IF cfld_.si_order.count > 0 THEN
            FOR si_ IN cfld_.si_order.first .. cfld_.si_order.last LOOP
               natr ('v', cfld_.si_order(si_), attrs_);
               Xml_Node (doc_, nd_si_, 's', attrs_); -- s for a string, which we assume, for now
            END LOOP;
         END IF;
      END LOOP;
      nd_el_ := Xml_Node (doc_, nd_pcd_, 'extLst');
      attrs_.delete;
      natr ('uri', Get_Guid, attrs_);
      attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
      nd_ex_ := Xml_Node (doc_, nd_el_, 'ext', attrs_);
      Xml_Node (doc_, nd_ex_, 'pivotCacheDefinition', 'x14');

      Add1Xml (excel_, rep('xl/pivotCache/pivotCacheDefinition:P1.xml',pc_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);


      -- xl/pivotCache/pivotCacheRecords:P1.xml
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
      attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
      attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
      attr ('mc:Ignorable', 'xr', attrs_);
      attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
      attr ('count', to_char(cache_.ds_range.br.r - cache_.ds_range.tl.r), attrs_);
      nd_pcd_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'pivotCacheRecords', attrs_);

      FOR r_ IN cache_.ds_range.tl.r+1 .. cache_.ds_range.br.r LOOP
         nd_row_ := Xml_Node (doc_, nd_pcd_, 'r');
         FOR c_ IN cache_.cf_order.first .. cache_.cf_order.last LOOP
            cfld_   := cache_.cached_fields(cache_.cf_order(c_));
            xl_col_ := cache_.ds_range.tl.c + c_ - 1; -- one based, not zero
            natr ('v', Get_Cell_Cache_Value (xl_col_, r_, sh_, cfld_.shared_items), attrs_);
            tag_    := Get_Cell_Cache_Tag (xl_col_, r_, sh_, cfld_.rollup_fn);
            Xml_Node (doc_, nd_row_, tag_, attrs_);
         END LOOP;
      END LOOP;
      Add1Xml (excel_, rep('xl/pivotCache/pivotCacheRecords:P1.xml',pc_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);


      -- One _rel file per pivot cache.  Technically, it could be possible for
      -- there to be multiple record-files per cache, but this won't happen in
      -- this program - at least for the time being.
      -- xl/pivotCache/_rels/pivotCacheDefinition:P1.xml.rels
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
      nd_rels_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

      natr ('Id', 'rId1', attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords', attrs_);
      attr ('Target', rep ('pivotCacheRecords:P1.xml', pc_), attrs_);
      Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);

      Add1Xml (excel_, rep('xl/pivotCache/_rels/pivotCacheDefinition:P1.xml.rels',pc_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);

   END LOOP;

END Finish_Pivot_Caches;

FUNCTION Combine_Arrays (
   arr1_ IN tp_col_filters,
   arr2_ IN tp_col_filters ) RETURN tp_col_filters
IS
   ix_  PLS_INTEGER;
   ret_ tp_col_filters;
BEGIN
   ix_ := arr1_.first;
   WHILE ix_ IS NOT null LOOP
      ret_(ix_) := arr1_(ix_);
      ix_ := arr1_.next(ix_);
   END LOOP;
   ix_ := arr2_.first;
   WHILE ix_ IS NOT null LOOP
      IF ret_.exists(ix_) THEN
         Raise_App_Error ('Duplicate indexes :P1 when combining arrays', to_char(ix_));
      END IF;
      ret_(ix_) := arr2_(ix_);
      ix_ := arr2_.next(ix_);
   END LOOP;
   RETURN ret_;
END Combine_Arrays;

PROCEDURE Check_Cell_Not_Exist (
   sh_  IN PLS_INTEGER,
   col_ IN PLS_INTEGER,
   row_ IN PLS_INTEGER )
IS BEGIN
   IF wb_.sheets(sh_).rows.exists(row_) AND wb_.sheets(sh_).rows(row_).exists(col_) THEN
      Raise_App_Error (
         'Pivot table has expaned into sheet/cell [:P1/:P2] which already contains data.' ||
         '  To avoid any chance of recursive references, this is not allowed.',
         sh_, Alfan_Cell (col_, row_)
      );
   END IF;
END Check_Cell_Not_Exist;


-----
-- Unravel_Json_To_Sheet()
--   Once rolled up, we first need to use the JSON to populate this workbook's
--   sheets.  Dependant on the pivot caches, it may well be that the new pivot
--   tables overwirte existing cells, which should not be allowed.
--   There are 2 versions of this function, the second is the initiator, while
--   the first is the recursor.
--
PROCEDURE Unravel_Json_To_Sheet (
   sh_       IN PLS_INTEGER,
   j_node_   IN json_object_t,
   init_col_ IN PLS_INTEGER,
   row_      IN OUT NOCOPY PLS_INTEGER )
IS
   col_      PLS_INTEGER;
   k_obj_    json_object_t;
   sum_obj_  json_object_t;
   keys_     json_key_list := j_node_.get_keys;
   s_keys_   json_key_list;
BEGIN

   FOR k_ IN keys_.first .. keys_.last LOOP

      col_ := init_col_;
      Check_Cell_Not_Exist (sh_, col_, row_);
      CellS (col_, row_, keys_(k_), sheet_ => sh_);

      k_obj_   := j_node_.get_object(keys_(k_));
      sum_obj_ := k_obj_.get_object('sum');
      s_keys_  := sum_obj_.get_keys;

      FOR l_ IN s_keys_.first .. s_keys_.last LOOP
         col_ := col_ + 1;
         Check_Cell_Not_Exist (sh_, col_, row_);
         CellN (col_, row_, sum_obj_.get_number(s_keys_(l_)), sheet_ => sh_);
      END LOOP;
      row_ := row_ + 1;

      IF k_obj_.has('breakdown') THEN
         Unravel_Json_To_Sheet (sh_, k_obj_.get_object('breakdown'), init_col_, row_);
      END IF;

   END LOOP;   

END Unravel_Json_To_Sheet;


PROCEDURE Unravel_Json_To_Sheet (
   pivot_id_ IN PLS_INTEGER,
   j_piv_    IN json_object_t )
IS
   aggs_arr_ tp_col_agg_fns := wb_.pivot_tables(pivot_id_).pivot_axes.col_agg_fns;
   ds_range_ tp_cell_range  := Get_Pivot_Source (pivot_id_);
   loc_      tp_cell_loc    := wb_.pivot_tables(pivot_id_).location_tl;
   sh_       PLS_INTEGER    := wb_.pivot_tables(pivot_id_).on_sheet;
   init_col_ PLS_INTEGER    := loc_.c;
   col_      PLS_INTEGER    := init_col_;
   row_      PLS_INTEGER    := loc_.r;
   desc_     VARCHAR2(2000);
   ix_       PLS_INTEGER;
   sum_obj_  json_object_t;
   keys_     json_key_list;
BEGIN

   CellS (col_, row_, 'Row Labels', sheet_ => sh_);
   ix_ := aggs_arr_.first;
   WHILE ix_ IS NOT null LOOP
      desc_ := CASE aggs_arr_(ix_)
         WHEN 'sum'   THEN 'Sum of '
         WHEN 'count' THEN 'Count of '
      END || Range_Col_Head_Name (ds_range_, ix_);
      col_ := col_ + 1;
      CellS (col_, row_, desc_, sheet_ => sh_);
      ix_ := aggs_arr_.next(ix_);
   END LOOP;
   row_ := row_ + 1;
   Unravel_Json_To_Sheet (sh_, j_piv_.get_object('breakdown'), init_col_, row_);

   col_ := init_col_;
   CellS (col_, row_, 'Grand Total', sheet_ => sh_);

   sum_obj_ := j_piv_.get_object('sum');
   keys_ := sum_obj_.get_keys;
   FOR k_ IN keys_.first .. keys_.last LOOP
      col_ := col_ + 1;
      CellN (col_, row_, sum_obj_.get_number(keys_(k_)), sheet_ => sh_);
   END LOOP;

END Unravel_Json_To_Sheet;

-----
-- Unravel_Json_Pivot_Table_Xml()
--   Once rolled up, we can use the JSON to build our PivotTable.xml series of
--   files.
PROCEDURE Unravel_Json_Pivot_Table_Xml (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   xml_nd_   IN            dbms_XmlDom.DomNode,
   j_node_   IN            json_object_t,
   cache_    IN OUT NOCOPY tp_pivot_cache,
   col_name_ IN            VARCHAR2 := null,
   si_val_   IN            VARCHAR2 := null )
IS
   nd_i_        dbms_XmlDom.DomNode;
   bkdn_obj_    json_object_t;
   keys_        json_key_list;
   attrs_       xml_attrs_arr;
   level_       PLS_INTEGER := j_node_.get_number('level');
   lv_          PLS_INTEGER := level_ - 1;
   v_           PLS_INTEGER := 0;
   si_val_loop_ VARCHAR2(32000);
BEGIN

   IF level_ > 0 THEN

      IF lv_ > 0 THEN
         attr ('r', to_char(lv_), attrs_);
      END IF;
      nd_i_ := Xml_Node (doc_, xml_nd_, 'i', attrs_);

      si_val_loop_ := cache_.cached_fields(col_name_).shared_items.first;
      WHILE si_val_loop_ IS NOT null LOOP
         EXIT WHEN si_val_loop_ = si_val_;
         v_ := v_ + 1;
         si_val_loop_ := cache_.cached_fields(col_name_).shared_items.next(si_val_loop_);
      END LOOP;

      attrs_.delete;
      IF v_ > 0 THEN
         attr ('v', v_, attrs_);
      END IF;
      Xml_Node (doc_, nd_i_, 'x', attrs_);

   END IF;

   IF j_node_.has('breakdown') THEN
      bkdn_obj_ := j_node_.get_object ('breakdown');
      keys_     := bkdn_obj_.get_keys;
      FOR k_ IN keys_.first .. keys_.last LOOP
         Unravel_Json_Pivot_Table_Xml (
            doc_      => doc_,
            xml_nd_   => xml_nd_,
            j_node_   => bkdn_obj_.get_object(keys_(k_)),
            cache_    => cache_,
            col_name_ => j_node_.get_string('colDesc'),
            si_val_   => keys_(k_)
         );
      END LOOP;
   END IF;

   IF level_ = 0 THEN
      natr ('t', 'grand', attrs_);
      nd_i_ := Xml_Node (doc_, xml_nd_, 'i', attrs_);
      Xml_Node (doc_, nd_i_, 'x');
   END IF;

END Unravel_Json_Pivot_Table_Xml;


-----
-- Json_Aggregates_From_Filters()
--   Given a cell-range (where our base-data is to be found) plus some filters
--   upon which we would like that data to be sieved, we can build a matrix of
--   the pivot table modelled as a JSON object.  At this stage, that JSON data
--   isn't located in the Excel sheet, but its shape is important to build out
--   the PivotTableX.xml flie, and later to stamp the data into the sheet.
--
FUNCTION Json_Aggregates_From_Filters (
   pivot_id_      IN PLS_INTEGER,
   level_         IN PLS_INTEGER,
   filter_vals_   IN tp_col_filters := tp_col_filters(), -- should be of length `level_`
   extra_filters_ IN tp_col_filters := tp_col_filters() ) RETURN json_object_t
IS

   pivot_            tp_pivot_table := wb_.pivot_tables(pivot_id_);
   cache_            tp_pivot_cache := wb_.pivot_caches(pivot_.cache_id);
   range_            tp_cell_range  := cache_.ds_range;
   vrollup_          tp_pivot_cols  := pivot_.pivot_axes.vrollups;
   aggregates_       tp_col_agg_fns := pivot_.pivot_axes.col_agg_fns;
   filters_          tp_col_filters := Combine_Arrays (filter_vals_, extra_filters_);
   agg_offset_       PLS_INTEGER;
   fc_offset_        PLS_INTEGER;

   is_leaf_          BOOLEAN       := vrollup_.count = level_;
   results_obj_      json_object_t := json_object_t();
   sum_obj_          json_object_t := json_object_t();
   breakdown_obj_    json_object_t := json_object_t();
   child_obj_        json_object_t;
   sum_val_          NUMBER        := 0;
   rec_count_        PLS_INTEGER   := 0;
   direct_sub_recs_  PLS_INTEGER   := 0;
   accum_sub_recs_   PLS_INTEGER   := 0;

   rg_row_start_     PLS_INTEGER := range_.tl.r + 1;
   rg_row_end_       PLS_INTEGER := range_.br.r;
   rg_col_start_     PLS_INTEGER := range_.tl.c - 1;
   col_              PLS_INTEGER;
   keep_             BOOLEAN;

   next_filter_vals_ tp_col_filters  := filter_vals_;
   next_level_       PLS_INTEGER     := level_ + 1;
   next_colid_       PLS_INTEGER;
   next_col_name_    VARCHAR2(32000);
   shared_item_      VARCHAR2(32000);

BEGIN

   IF vrollup_.count = 0 THEN
      Raise_App_Error ('There must be at least 1 rollup in a Pivot Table!');
   END IF;

   -- initiate values
   results_obj_.put ('level', level_);
   results_obj_.put ('isLeaf', is_leaf_);
   results_obj_.put ('count', 0);

   agg_offset_ := aggregates_.first;
   WHILE agg_offset_ IS NOT null LOOP
      IF aggregates_(agg_offset_) = 'sum' THEN
         sum_obj_.put (to_char(agg_offset_), 0);
      END IF;
      agg_offset_ := aggregates_.next(agg_offset_);
   END LOOP;

   IF is_leaf_ THEN

      -- For leaf elements, we revert to the data-source already inserted into
      -- our Excel sheet.  We do our own filtering of this data to find out if
      -- each record is of interest to us or not, and then base our pivot math
      -- calculations on that.  Parent elements will base their aggregation on
      -- this data as well.

      FOR r_ IN rg_row_start_ .. rg_row_end_ LOOP -- loop on dataset rows
         keep_      := true;
         fc_offset_ := filters_.first;
         WHILE fc_offset_ IS NOT null AND keep_ LOOP -- loop on each search criteria
            col_ := rg_col_start_ + fc_offset_;
            keep_ := keep_ AND Get_Cell_Value_Raw (col_, r_, range_.sheet_id, false) = filters_(fc_offset_);
            fc_offset_ := filters_.next(fc_offset_);
         END LOOP;
         IF keep_ THEN
            agg_offset_ := aggregates_.first;
            rec_count_  := rec_count_ + 1;
            WHILE agg_offset_ IS NOT null LOOP -- loop on aggregation array to print the results out
               col_ := rg_col_start_ + agg_offset_;
               sum_val_ := sum_obj_.get_number(to_char(agg_offset_)) + Get_Cell_Value_Num (col_, r_, range_.sheet_id);
               sum_obj_.put (to_char(agg_offset_), sum_val_);
               agg_offset_ := aggregates_.next(agg_offset_);
            END LOOP;
         END IF;
      END LOOP;
      results_obj_.put ('count', rec_count_);
      results_obj_.put ('sum', sum_obj_);

   ELSE
      -- If this is not the leaf, we need to recurse down to the next level of
      -- rollup.  Aggregations are added up from leaf-nodes, in order to limit
      -- looping on the source data as much as possible

      next_colid_    := vrollup_(next_level_);
      next_col_name_ := Range_Col_Head_Name (cache_.ds_range, next_colid_);
      results_obj_.put ('colId',   next_colid_);
      results_obj_.put ('colDesc', next_col_name_);

      -- Loop on each "shared item" already built into the cache
      shared_item_ := cache_.cached_fields(next_col_name_).shared_items.first;
      WHILE shared_item_ IS NOT null LOOP

         next_filter_vals_(next_colid_) := shared_item_;
         child_obj_ := Json_Aggregates_From_Filters (
            pivot_id_      => pivot_id_,
            level_         => next_level_,
            filter_vals_   => next_filter_vals_, -- varchar index-by col_ix_, should be same length as `level_`
            extra_filters_ => extra_filters_
         );

         -- We only attach elements that have records in this filter criteria
         IF child_obj_.get_number('count') > 0 THEN

            breakdown_obj_.put (shared_item_, child_obj_);

            direct_sub_recs_  := direct_sub_recs_ + 1;
            IF child_obj_.has('accum-sub-records') THEN
               accum_sub_recs_ := accum_sub_recs_ + child_obj_.get_number('accum-sub-records');
            END IF;
            rec_count_ := rec_count_ + child_obj_.get_number('count');
            agg_offset_ := aggregates_.first;
            WHILE agg_offset_ IS NOT null LOOP
               sum_val_ := sum_obj_.get_number(to_char(agg_offset_));
               sum_val_ := sum_val_ + child_obj_.get_object('sum').get_number(to_char(agg_offset_));
               sum_obj_.put (to_char(agg_offset_), sum_val_);
               agg_offset_ := aggregates_.next(agg_offset_);
            END LOOP;
         END IF;

         shared_item_ := cache_.cached_fields(next_col_name_).shared_items.next(shared_item_);
      END LOOP;

      results_obj_.put ('direct-sub-records', direct_sub_recs_);
      results_obj_.put ('accum-sub-records', accum_sub_recs_ + direct_sub_recs_);
      results_obj_.put ('count', rec_count_);
      results_obj_.put ('sum', sum_obj_);
      results_obj_.put ('breakdown', breakdown_obj_);

   END IF;

   IF level_ = 0 THEN
      Unravel_Json_To_Sheet (pivot_id_, results_obj_);
   END IF;
   RETURN results_obj_;

END Json_Aggregates_From_Filters;


-----
-- Finish_Pivot_Tables()
--   Must be called after Build_Pivot_Caches(), which isn't hard to achieve on
--   account of being called at the start of the Finish() process.  It's worth
--   making a note of anyway to avoid potential problem in future.
--
PROCEDURE Finish_Pivot_Tables (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_         dbms_XmlDom.DomDocument;
   attrs_       xml_attrs_arr;
   nd_ptd_      dbms_XmlDom.DomNode;
   nd_pfs_      dbms_XmlDom.DomNode;
   nd_pf_       dbms_XmlDom.DomNode;
   nd_is_       dbms_XmlDom.DomNode;
   nd_ri_       dbms_XmlDom.DomNode;
   nd_exl_      dbms_XmlDom.DomNode;
   nd_ext_      dbms_XmlDom.DomNode;
   nd_rels_     dbms_XmlDom.DomNode;
   nd_dfs_      dbms_XmlDom.DomNode;
   j_piv_       json_object_t;
   pt_region_   tp_cell_range;
   cache_       tp_pivot_cache;
   cf_          tp_cache_field;
   shared_item_ VARCHAR2(32000);
   agg_col_     PLS_INTEGER;
   prefix_      VARCHAR2(50);

BEGIN

   FOR pt_ IN 1 .. wb_.pivot_tables.count LOOP

      j_piv_ := Json_Aggregates_From_Filters (pivot_id_ => pt_, level_ => 0);
      Trace ('Pivot table :P1: ' || j_piv_.stringify);

      cache_ := wb_.pivot_caches(wb_.pivot_tables(pt_).cache_id);

      -- xl/pivotTables/pivotTable:P1.xml
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
      attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
      attr ('mc:Ignorable', 'xr', attrs_);
      attr ('name', wb_.pivot_tables(pt_).pivot_name, attrs_);
      attr ('cacheId', wb_.pivot_tables(pt_).cache_id, attrs_);
      attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
      attr ('xr:uid', Get_Guid, attrs_);
      attr ('applyNumberFormats', '0', attrs_);
      attr ('applyBorderFormats', '0', attrs_);
      attr ('applyFontFormats', '0', attrs_);
      attr ('applyPatternFormats', '0', attrs_);
      attr ('applyAlignmentFormats', '0', attrs_);
      attr ('applyWidthHeightFormats', '1', attrs_);
      attr ('dataCaption', 'Values', attrs_);
      attr ('createdVersion', '7', attrs_); -- Version of Excel in which this pivot was created!
      attr ('updatedVersion', '7', attrs_);
      attr ('minRefreshableVersion', '3', attrs_); -- Minimum version of Excel which is compatible (apparently)
      attr ('useAutoFormatting', '1', attrs_);
      attr ('itemPrintTitles', '1', attrs_);
      attr ('indent', '0', attrs_);
      attr ('outline', '1', attrs_);
      attr ('outlineData', '1', attrs_);
      attr ('multipleFieldFilters', '0', attrs_);
      nd_ptd_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'pivotTableDefinition', attrs_);

      wb_.pivot_tables(pt_).pivot_height := j_piv_.get_number('accum-sub-records') + 2; -- +header, +grand-totals rows
      wb_.pivot_tables(pt_).pivot_width  := wb_.pivot_tables(pt_).pivot_axes.col_agg_fns.count + 1; -- assume for now that we only do v-rollups
      pt_region_ := tp_cell_range (
         sheet_id => wb_.pivot_tables(pt_).on_sheet,
         tl       => wb_.pivot_tables(pt_).location_tl,
         br       => tp_cell_loc (
            c => wb_.pivot_tables(pt_).location_tl.c + wb_.pivot_tables(pt_).pivot_width - 1,
            r => wb_.pivot_tables(pt_).location_tl.r + wb_.pivot_tables(pt_).pivot_height - 1
         )
      );

      natr ('ref', Alfan_Range(pt_region_), attrs_);
      attr ('firstHeaderRow', '1', attrs_);
      attr ('firstDataRow', '1', attrs_);
      attr ('firstDataCol', '1', attrs_);
      Xml_Node (doc_, nd_ptd_, 'location', attrs_);

      natr ('count', to_char(cache_.cf_order.count), attrs_);
      nd_pfs_ := Xml_Node (doc_, nd_ptd_, 'pivotFields', attrs_);

      FOR cf_ix_ IN cache_.cf_order.first .. cache_.cf_order.last LOOP

         -- the rollup in the pivot-table needn't be the same as the rollup in
         -- the cache because a cache could serve multiple tables.  The rollup
         -- calculated here must therefore be that taken from the table
         attrs_.delete;
         CASE Get_Agg_Fn_From_Axes (wb_.pivot_tables(pt_).pivot_axes, cf_ix_)
            WHEN 'row' THEN attr ('axis', 'axisRow', attrs_);
            WHEN 'sum' THEN attr ('dataField', '1', attrs_);
            ELSE null;
         END CASE;
         attr ('showAll', 0, attrs_);
         nd_pf_ := Xml_Node (doc_, nd_pfs_, 'pivotField', attrs_);

         cf_ := cache_.cached_fields(cache_.cf_order(cf_ix_));
         IF cf_.shared_items.count > 0 THEN
            natr ('count', to_char(cf_.shared_items.count + 1), attrs_);
            nd_is_ := Xml_Node (doc_, nd_pf_, 'items', attrs_);

            shared_item_ := cf_.shared_items.first;
            WHILE shared_item_ IS NOT null LOOP
               natr ('x', cf_.shared_items(shared_item_), attrs_);
               Xml_Node (doc_, nd_is_, 'item', attrs_);
               shared_item_ := cf_.shared_items.next(shared_item_);
            END LOOP;
            natr ('t', 'default', attrs_);
            Xml_Node (doc_, nd_is_, 'item', attrs_);
         END IF;

      END LOOP;

      natr ('count', to_char(wb_.pivot_tables(pt_).pivot_axes.vrollups.count), attrs_);
      nd_pfs_ := Xml_Node (doc_, nd_ptd_, 'rowFields', attrs_);

      FOR r_ IN 1 .. wb_.pivot_tables(pt_).pivot_axes.vrollups.count LOOP
         natr ('x', to_char(wb_.pivot_tables(pt_).pivot_axes.vrollups(r_) - 1), attrs_);
         Xml_Node (doc_, nd_pfs_, 'field', attrs_);
      END LOOP;

      natr ('count', to_char(j_piv_.get_number('accum-sub-records') + 1), attrs_);
      nd_ri_ := Xml_Node (doc_, nd_ptd_, 'rowItems', attrs_);

      Unravel_Json_Pivot_Table_Xml (doc_, nd_ri_, j_piv_, cache_);

      natr ('count', 1, attrs_);
      Xml_Node (doc_, Xml_Node (doc_, nd_ptd_, 'colItems', attrs_), 'i');

      natr ('count', to_char(wb_.pivot_tables(pt_).pivot_axes.col_agg_fns.count), attrs_);
      nd_dfs_ := Xml_Node (doc_, nd_ptd_, 'dataFields', attrs_);

      agg_col_ := wb_.pivot_tables(pt_).pivot_axes.col_agg_fns.first;
      WHILE agg_col_ IS NOT null LOOP
         prefix_ := CASE wb_.pivot_tables(pt_).pivot_axes.col_agg_fns(agg_col_)
            WHEN 'sum' THEN 'Sum of '
         END;
         natr ('name', prefix_ || Range_Col_Head_Name (cache_.ds_range, agg_col_), attrs_);
         attr ('fld', agg_col_ - 1, attrs_); -- zero based, I think
         attr ('baseField', '0', attrs_); -- used with showDataAs, which we aren't using for now
         attr ('baseItem', '0', attrs_);
         Xml_Node (doc_, nd_dfs_, 'dataField', attrs_);
         agg_col_ := wb_.pivot_tables(pt_).pivot_axes.col_agg_fns.next(agg_col_);
      END LOOP;

      natr ('name', 'PivotStyleLight16', attrs_);
      attr ('showRowHeaders', '1', attrs_);
      attr ('showColHeaders', '1', attrs_);
      attr ('showRowStripes', '0', attrs_);
      attr ('showColStripes', '0', attrs_);
      attr ('showLastColumn', '1', attrs_);
      Xml_Node (doc_, nd_ptd_, 'pivotTableStyleInfo', attrs_);

      nd_exl_ := Xml_Node (doc_, nd_ptd_, 'extLst');

      natr ('uri', Get_Guid, attrs_);
      attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
      nd_ext_ := Xml_Node (doc_, nd_exl_, 'ext', attrs_);

      natr ('hideValuesRow', '1', attrs_);
      attr ('xmlns:xm', 'http://schemas.microsoft.com/office/excel/2006/main', attrs_);
      Xml_Node (doc_, nd_ext_, 'pivotTableDefinition', 'x14', attrs_);

      natr ('uri', Get_Guid, attrs_);
      attr ('xmlns:xpdl', 'http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout', attrs_);
      nd_ext_ := Xml_Node (doc_, nd_exl_, 'ext', attrs_);
      Xml_Node (doc_, nd_ext_, 'pivotTableDefinition16', 'xpdl');

      Add1Xml (excel_, rep('xl/pivotTables/pivotTable:P1.xml',pt_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);


      -- One _rel file per pivot table.
      -- xl/pivotTables/_rels/pivotTable:P1.xml.rels
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
      nd_rels_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

      natr ('Id', 'rId1', attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition', attrs_);
      attr ('Target', rep ('../pivotCache/pivotCacheDefinition:P1.xml', to_char(wb_.pivot_tables(pt_).cache_id)), attrs_);
      Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);

      Add1Xml (excel_, rep('xl/pivotTables/_rels/pivotTable:P1.xml.rels',pt_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);

   END LOOP;
END Finish_Pivot_Tables;


PROCEDURE Finish_Drawings_Rels (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_     dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   attrs_   xml_attrs_arr;
   nd_rels_ dbms_XmlDom.DomNode;
BEGIN

   IF wb_.images.count = 0 THEN
      goto skip_drawings_rels;
   END IF;

   -- xl/drawings/_rels/drawing1.xml.rels
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rels_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

   FOR dr_ IN 1 .. wb_.images.count LOOP
      natr ('Id', 'rId' || dr_, attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', attrs_);
      attr ('Target', rep ('../media/image:P1.:P2', dr_, wb_.images(dr_).extension), attrs_);
      Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      Add1File (
         zipped_blob_ => excel_,
         filename_    => rep ('xl/media/image:P1.:P2', dr_, wb_.images(dr_).extension),
         content_     => wb_.images(dr_).img_blob
      );
   END LOOP;

   Add1Xml (excel_, 'xl/drawings/_rels/drawing1.xml.rels', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

   <<skip_drawings_rels>>
   null;

END Finish_Drawings_Rels;

PROCEDURE Finish_Worksheet (
   excel_ IN OUT NOCOPY BLOB,
   s_     IN            PLS_INTEGER )
IS
   doc_     dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   attrs_   xml_attrs_arr;
   nd_ws_   dbms_XmlDom.DomNode;
   nd_svs_  dbms_XmlDom.DomNode;
   nd_sv_   dbms_XmlDom.DomNode;
   nd_cls_  dbms_XmlDom.DomNode;
   nd_sd_   dbms_XmlDom.DomNode;
   nd_r_    dbms_XmlDom.DomNode;
   nd_c_    dbms_XmlDom.DomNode;
   nd_mc_   dbms_XmlDom.DomNode;
   nd_dvs_  dbms_XmlDom.DomNode;
   nd_dv_   dbms_XmlDom.DomNode;
   nd_h_    dbms_XmlDom.DomNode;
   row_     PLS_INTEGER := wb_.sheets(s_).rows.first;
   col_     PLS_INTEGER;
   col_min_ PLS_INTEGER := 16384;
   col_max_ PLS_INTEGER := 1;
   id_      PLS_INTEGER := 1;
BEGIN

   WHILE row_ IS NOT null LOOP
      col_min_ := least (col_min_, wb_.sheets(s_).rows(row_).first);
      col_max_ := greatest (col_max_, wb_.sheets(s_).rows(row_).last);
      row_  := wb_.sheets(s_).rows.next(row_);
   END LOOP;

   -- xl/worksheets/sheet:P1.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
   attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
   attr ('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', attrs_);
   attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
   --attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
   attr ('mc:Ignorable', 'x14ac', attrs_);
   attr ('xr:uid', Get_Guid, attrs_);
   nd_ws_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'worksheet', attrs_);
   IF wb_.sheets(s_).tabcolor IS NOT null THEN
      natr ('rgb', wb_.sheets(s_).tabcolor, attrs_);
      Xml_Node (doc_, Xml_Node(doc_,nd_ws_,'sheetPr'), 'tabColor', attrs_);
   END IF;

   natr (
      'ref', Alfan_Range (
         col_tl_ => col_min_, row_tl_ => wb_.sheets(s_).rows.first,
         col_br_ => col_max_, row_br_ => wb_.sheets(s_).rows.last
      ), attrs_
   );
   Xml_Node (doc_, nd_ws_, 'dimension', attrs_);

   nd_svs_ := Xml_Node (doc_, nd_ws_, 'sheetViews');
   attrs_.delete;
   IF s_ = 1 THEN attr ('tabSelected', '1', attrs_); END IF;
   attr ('workbookViewId', '0', attrs_);
   nd_sv_  := Xml_Node (doc_, nd_svs_, 'sheetView', attrs_);

   IF wb_.sheets(s_).freeze_rows + wb_.sheets(s_).freeze_cols > 0 THEN
      natr ('activePane', 'bottomLeft', attrs_);
      attr ('state', 'frozen', attrs_);
      IF wb_.sheets(s_).freeze_rows > 0 AND wb_.sheets(s_).freeze_cols > 0 THEN
         attr ('xSplit', wb_.sheets(s_).freeze_cols, attrs_);
         attr ('ySplit', wb_.sheets(s_).freeze_rows, attrs_);
         attr ('topLeftCell', Alfan_Cell (wb_.sheets(s_).freeze_cols+1, wb_.sheets(s_).freeze_rows+1), attrs_);
      ELSIF wb_.sheets(s_).freeze_rows > 0 THEN
         attr ('ySplit', wb_.sheets(s_).freeze_rows, attrs_);
         attr ('topLeftCell', Alfan_Cell (1, wb_.sheets(s_).freeze_rows+1), attrs_);
      ELSIF wb_.sheets(s_).freeze_cols > 0 THEN
         attr ('xSplit', wb_.sheets(s_).freeze_cols, attrs_);
         attr ('topLeftCell', Alfan_Cell (wb_.sheets(s_).freeze_cols+1, 1), attrs_);
      END IF;
      Xml_Node (doc_, nd_sv_, 'pane', attrs_);
   ELSE
      natr ('activeCell', 'A1', attrs_);
      attr ('sqref', 'A1', attrs_);
      Xml_Node (doc_, nd_sv_, 'selection', attrs_);
   END IF;

   attrs_.delete;
   natr ('defaultRowHeight', '15', attrs_);
   attr ('x14ac:dyDescent', '0.25', attrs_);
   Xml_Node (doc_, nd_ws_, 'sheetFormatPr', attrs_);

   IF wb_.sheets(s_).widths.count > 0 THEN
      nd_cls_ := Xml_Node (doc_, nd_ws_, 'cols');
      attrs_.delete;
      col_ := wb_.sheets(s_).widths.first;
      WHILE col_ IS NOT null LOOP
         natr ('min', col_, attrs_);
         attr ('max', col_, attrs_);
         attr ('width', to_char (wb_.sheets(s_).widths(col_), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'), attrs_);
         attr ('customWidth', '1', attrs_);
         Xml_Node (doc_, nd_cls_, 'col', attrs_);
         col_ := wb_.sheets(s_).widths.next(col_);
      END LOOP;
   END IF;

   nd_sd_ := Xml_Node (doc_, nd_ws_, 'sheetData');
   row_   := wb_.sheets(s_).rows.first;
   WHILE row_ IS NOT null LOOP
      natr ('r', to_char(row_), attrs_);
      attr ('spans', to_char(col_min_) || ':' || to_char(col_max_), attrs_);
      IF wb_.sheets(s_).row_fmts.exists(row_) AND wb_.sheets(s_).row_fmts(row_).height IS NOT null THEN
         attr ('customHeight', '1', attrs_);
         attr ('ht', to_char (wb_.sheets(s_).row_fmts(row_).height, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'), attrs_);
      END IF;
      nd_r_ := Xml_Node (doc_, nd_sd_, 'row', attrs_);

      col_ := wb_.sheets(s_).rows(row_).first;
      WHILE col_ IS NOT null LOOP
         natr ('r', Alfan_Cell (col_, row_), attrs_);
         IF wb_.sheets(s_).rows(row_)(col_).datatype IN (CELL_DT_STRING_, CELL_DT_HYPERLINK_) THEN
            attr ('t', 's', attrs_);
         END IF;
         IF wb_.sheets(s_).rows(row_)(col_).style IS NOT null THEN
            attr ('s', to_char(wb_.sheets(s_).rows(row_)(col_).style), attrs_);
         END IF;
         nd_c_ := Xml_Node (doc_, nd_r_, 'c', attrs_);
         IF wb_.sheets(s_).rows(row_)(col_).formula_idx IS NOT null THEN
            Xml_Text_Node (doc_, nd_c_, 'f', wb_.formulas(wb_.sheets(s_).rows(row_)(col_).formula_idx));
         END IF;
         Xml_Text_Node (doc_, nd_c_, 'v', to_char(wb_.sheets(s_).rows(row_)(col_).value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'));
         col_ := wb_.sheets(s_).rows(row_).next(col_);
      END LOOP;
      row_ := wb_.sheets(s_).rows.next(row_);
   END LOOP;

   FOR af_ IN 1 .. wb_.sheets(s_).autofilters.count LOOP
      attrs_.delete;
      natr (
         'ref', Alfan_Range (
            col_tl_ => nvl (wb_.sheets(s_).autofilters(af_).column_start, col_min_),
            row_tl_ => nvl (wb_.sheets(s_).autofilters(af_).row_start, wb_.sheets(s_).rows.first),
            col_br_ => coalesce (wb_.sheets(s_).autofilters(af_).column_end, wb_.sheets(s_).autofilters(af_).column_start, col_max_),
            row_br_ => nvl (wb_.sheets(s_).autofilters(af_).row_end, wb_.sheets(s_).rows.last)
         ), attrs_
      );
      Xml_Node (doc_, nd_ws_, 'autoFilter', attrs_);
   END LOOP;

   IF wb_.sheets(s_).mergecells.count > 0 THEN
      natr ('count', to_char(wb_.sheets(s_).mergecells.count), attrs_);
      nd_mc_ := Xml_Node (doc_, nd_ws_, 'mergeCells', attrs_);
      FOR mg_ IN 1 .. wb_.sheets(s_).mergecells.count LOOP
         natr ('ref', wb_.sheets(s_).mergecells(mg_), attrs_);
         Xml_Node (doc_, nd_mc_, 'mergeCell', attrs_);
      END LOOP;
   END IF;

   IF wb_.sheets(s_).validations.count > 0 THEN
      natr ('count', to_char(wb_.sheets(s_).validations.count), attrs_);
      nd_dvs_ := Xml_Node (doc_, nd_ws_, 'dataValidations');

      FOR v_ IN wb_.sheets(s_).validations.count LOOP
         attrs_.delete;
         natr ('type', wb_.sheets(s_).validations(v_).type, attrs_);
         attr ('errorStyle', wb_.sheets(s_).validations(v_).errorstyle, attrs_);
         attr ('allowBlank', CASE WHEN nvl(wb_.sheets(s_).validations(v_).allowBlank, true) THEN '1' ELSE '0' END, attrs_);
         attr ('sqref', wb_.sheets(s_).validations(v_).sqref, attrs_);
         IF wb_.sheets(s_).validations(v_).prompt IS NOT null THEN
            attr ('showInputMessage', '1', attrs_);
            attr ('prompt', wb_.sheets(s_).validations(v_).prompt, attrs_);
            IF wb_.sheets(s_).validations(v_).title IS NOT null THEN
               attr ('promptTitle', wb_.sheets(s_).validations(v_).title, attrs_);
            END IF;
         END IF;
         IF wb_.sheets(s_).validations(v_).showerrormessage THEN
            attr ('showErrorMessage', '1', attrs_);
            IF wb_.sheets(s_).validations(v_).error_title IS NOT null THEN
               attr ('errorTitle', wb_.sheets(s_).validations(v_).error_title, attrs_);
            END IF;
            IF wb_.sheets(s_).validations(v_).error_txt IS NOT null THEN
               attr ('error', wb_.sheets(s_).validations(v_).error_txt, attrs_);
            END IF;
         END IF;
         nd_dv_ := Xml_Node (doc_, nd_dvs_, 'dataValidation', attrs_);

         IF wb_.sheets(s_).validations(v_).formula1 IS NOT null THEN
            Xml_Text_Node (doc_, nd_dv_, 'formula1', wb_.sheets(s_).validations(v_).formula1);
         END IF;
         IF wb_.sheets(s_).validations(v_).formula2 IS NOT null THEN
            Xml_Text_Node (doc_, nd_dv_, 'formula2', wb_.sheets(s_).validations(v_).formula2);
         END IF;
      END LOOP;
   END IF;

   IF wb_.sheets(s_).hyperlinks.count > 0 THEN
      nd_h_ := Xml_Node (doc_, nd_ws_, 'hyperlinks');
      FOR h_ IN 1 .. wb_.sheets(s_).hyperlinks.count LOOP
         natr ('ref', wb_.sheets(s_).hyperlinks(h_).cell, attrs_);
         attr ('r:id', rep ('rId:P1', id_), attrs_);
         Xml_Node (doc_, nd_h_, 'hyperlink', attrs_);
         id_ := id_ + 1;
      END LOOP;
   END IF;

   -- pivot tables need to be inserted here
   -- Question about whether pivot tables need to be generated before or after
   -- images (drawings) and comments.  Assuming that the designer of the Excel
   -- document is careful, there shouldn't be too many overlaps of normal data
   -- with pivoted data (which can expand quite easily to cover large areas of
   -- a sheet).  Images and comments should still be allowed to be placed over
   -- pivot tables though (at least in principlet)

   natr ('left', '0.7', attrs_);
   attr ('right', '0.7', attrs_);
   attr ('top', '0.75', attrs_);
   attr ('bottom', '0.75', attrs_);
   attr ('header', '0.3', attrs_);
   attr ('footer', '0.3', attrs_);
   Xml_Node (doc_, nd_ws_, 'pageMargins', attrs_);

   IF wb_.sheets(s_).drawings.count > 0 THEN
      natr ('r:id', rep ('rId:P1', id_), attrs_);
      Xml_Node (doc_, nd_ws_, 'drawing', attrs_);
      id_ := id_ + 1;
   END IF;
   
   IF wb_.sheets(s_).comments.count > 0 THEN
      natr ('r:id', 'rId' || id_, attrs_);
      Xml_Node (doc_, nd_ws_, 'legacyDrawing', attrs_);
   END IF;

   Add1Xml (excel_, rep('xl/worksheets/sheet:P1.xml',to_char(s_)), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Worksheet;

PROCEDURE Finish_Ws_Relationships (
   excel_ IN OUT NOCOPY BLOB,
   s_     IN            PLS_INTEGER )
IS
   id_            PLS_INTEGER := 1;
   nr_hyperlinks_ PLS_INTEGER := wb_.sheets(s_).hyperlinks.count;
   nr_comments_   PLS_INTEGER := wb_.sheets(s_).comments.count;
   nr_pivots_     PLS_INTEGER := wb_.sheets(s_).pivots_list.count;
   nr_drawings_   PLS_INTEGER := wb_.sheets(s_).drawings.count;
   pivot_id_      PLS_INTEGER;
   doc_           dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   attrs_         xml_attrs_arr;
   nd_rels_       dbms_XmlDom.DomNode;
BEGIN

   IF nr_hyperlinks_ = 0 AND nr_comments_ = 0 AND nr_pivots_ = 0 AND nr_drawings_ = 0 THEN
      goto skip_relationships;
   END IF;

   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rels_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

   FOR h_ IN 1 .. nr_hyperlinks_ LOOP
      IF wb_.sheets(s_).hyperlinks(h_).url IS NOT null THEN
         natr ('Id', rep ('rId:P1', id_), attrs_);
         attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', attrs_);
         attr ('Target',  wb_.sheets(s_).hyperlinks(h_).url, attrs_);
         attr ('TargetMode', 'External', attrs_);
         Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
         id_ := id_ + 1;
      END IF;
   END LOOP;
   IF nr_drawings_ > 0 THEN
      natr ('Id', rep ('rId:P1', id_), attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', attrs_);
      attr ('Target', rep ('../drawings/drawing:P1.xml', to_char(s_)), attrs_);
      Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      id_ := id_ + 1;
   END IF;
   IF nr_comments_ > 0 THEN
      natr ('Id', rep ('rId:P1', id_), attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing', attrs_);
      attr ('Target', rep ('../drawings/vmlDrawing:P1.vml', to_char(s_)), attrs_);
      Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      id_ := id_ + 1;
      natr ('Id', rep('rId:P1', id_), attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', attrs_);
      attr ('Target', rep ('../comments:P1.xml', to_char(s_)), attrs_);
      Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      id_ := id_ + 1;
   END IF;
   FOR spid_ IN 1 .. wb_.sheets(s_).pivots_list.count LOOP
      pivot_id_ := wb_.sheets(s_).pivots_list(spid_);
      natr ('Id', 'rId' || to_char(id_), attrs_);
      attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable', attrs_);
      attr ('Target', rep ('../pivotTables/pivotTable:P1.xml', pivot_id_), attrs_);
      Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      id_ := id_ + 1;
   END LOOP;

   Add1Xml (excel_, rep('xl/worksheets/_rels/sheet:P1.xml.rels',to_char(s_)), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

   <<skip_relationships>>
   null;

END Finish_Ws_Relationships;


PROCEDURE Calc_Image_Col_And_Row (
   col_      IN OUT NOCOPY PLS_INTEGER,
   row_      IN OUT NOCOPY PLS_INTEGER,
   col_offs_ IN OUT NOCOPY NUMBER,
   row_offs_ IN OUT NOCOPY NUMBER,
   drawing_  IN            tp_drawing,
   s_        IN            PLS_INTEGER )
IS
   scale_          NUMBER := nvl (drawing_.scale, 1);
   img_width_      NUMBER := wb_.images(drawing_.img_id).width  * scale_;
   img_height_     NUMBER := wb_.images(drawing_.img_id).height * scale_;
   img_width_rem_  NUMBER := img_width_;
   img_height_rem_ NUMBER := img_height_;
   img_colspan_    PLS_INTEGER;
   img_rowspan_    PLS_INTEGER;
   col_width_      NUMBER;
   row_height_     NUMBER;
BEGIN
   IF wb_.sheets(s_).widths.count = 0 THEN
      -- If no widths have been set, we can assume that all columns are set to
      -- the default widths => 64 px = 1 col = 609600
      img_colspan_ := trunc (img_width_/64);
      col_         := drawing_.col - 1 + img_colspan_;
      col_offs_    := trunc((img_width_-img_colspan_*64)*9525);
   ELSE
      col_ := drawing_.col;
      LOOP
         col_width_ := CASE
            WHEN not wb_.sheets(s_).widths.exists(col_) THEN 64
            ELSE round(7*wb_.sheets(s_).widths(col_))
         END;
         EXIT WHEN img_width_rem_ < col_width_;
         img_width_rem_ := img_width_rem_ - col_width_;
         col_ := col_ + 1;
      END LOOP;
      col_ := col_ - 1;
      col_offs_ := trunc(img_width_rem_ * 9525);
   END IF;
   IF wb_.sheets(s_).row_fmts.count = 0 THEN
      -- If no heights have been set then we assume the default row heights of
      -- => 20 px = 1 row = 190500
      img_rowspan_ := trunc (img_height_/20);
      row_         := drawing_.row - 1 + img_rowspan_;
      row_offs_    := trunc((img_height_- img_rowspan_*20) * 9525);
   ELSE
      row_ := drawing_.row;
      LOOP
         row_height_ := CASE
            WHEN wb_.sheets(s_).row_fmts.exists(row_) AND wb_.sheets(s_).row_fmts(row_).height IS NOT null THEN
               round (4 * wb_.sheets(s_).row_fmts(row_).height / 3)
            ELSE 20
         END;
         EXIT WHEN img_height_rem_ < row_height_;
         img_height_rem_ := img_height_rem_ - row_height_;
         row_ := row_ + 1;
      END LOOP;
      row_offs_ := trunc(img_height_rem_ * 9525);
      row_ := row_ - 1;
   END IF;
END Calc_Image_Col_And_Row;

PROCEDURE Finish_Ws_Drawings (
   excel_ IN OUT NOCOPY BLOB,
   s_     IN            PLS_INTEGER )
IS
   doc_        dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_ws_      dbms_XmlDom.DomNode;
   nd_tc_      dbms_XmlDom.DomNode;
   nd_fr_      dbms_XmlDom.DomNode;
   nd_to_      dbms_XmlDom.DomNode;
   nd_pi_      dbms_XmlDom.DomNode;
   nd_nv_      dbms_XmlDom.DomNode;
   nd_cn_      dbms_XmlDom.DomNode;
   nd_bf_      dbms_XmlDom.DomNode;
   nd_bl_      dbms_XmlDom.DomNode;
   nd_el_      dbms_XmlDom.DomNode;
   nd_et_      dbms_XmlDom.DomNode;
   attrs_      xml_attrs_arr;
   drawing_    tp_drawing;
   to_col_     PLS_INTEGER;
   to_row_     PLS_INTEGER;
   col_ovfl_  NUMBER;
   row_ovfl_   NUMBER;
BEGIN

   IF wb_.sheets(s_).drawings.count = 0 THEN
      goto skip_drawings;
   END IF;

   -- xl/drawings/drawing:P1.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   natr ('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', attrs_);
   attr ('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main', attrs_);
   nd_ws_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'wsDr', 'xdr', attrs_);

   FOR img_ IN 1 .. wb_.sheets(s_).drawings.count LOOP

      drawing_ := wb_.sheets(s_).drawings(img_);
      Calc_Image_Col_And_Row (to_col_, to_row_, col_ovfl_, row_ovfl_, drawing_, s_);

      natr ('editAs', 'oneCell', attrs_);
      nd_tc_ := Xml_Node (doc_, nd_ws_, 'twoCellAnchor', 'xdr', attrs_);

      nd_fr_ := Xml_Node (doc_, nd_tc_, 'from', 'xdr');
      Xml_Text_Node (doc_, nd_fr_, 'col', to_char(drawing_.col-1), 'xdr');
      Xml_Text_Node (doc_, nd_fr_, 'colOff', '0', 'xdr');
      Xml_Text_Node (doc_, nd_fr_, 'row', to_char(drawing_.row-1), 'xdr');
      Xml_Text_Node (doc_, nd_fr_, 'rowOff', '0', 'xdr');

      nd_to_ := Xml_Node (doc_, nd_tc_, 'to', 'xdr');
      Xml_Text_Node (doc_, nd_to_, 'col', to_char(to_col_), 'xdr');
      Xml_Text_Node (doc_, nd_to_, 'colOff', to_char(col_ovfl_), 'xdr');
      Xml_Text_Node (doc_, nd_to_, 'row', to_char(to_row_), 'xdr');
      Xml_Text_Node (doc_, nd_to_, 'rowOff', to_char(row_ovfl_), 'xdr');

      nd_pi_ := Xml_Node (doc_, nd_tc_, 'pic', 'xdr');
      nd_nv_ := Xml_Node (doc_, nd_pi_, 'nvPicPr', 'xdr');

      natr ('id', '3', attrs_);
      attr ('name', coalesce (drawing_.name, 'Picture '||img_), attrs_);
      IF drawing_.title       IS NOT null THEN attr('title', drawing_.title, attrs_); END IF;
      IF drawing_.description IS NOT null THEN attr('descr', drawing_.description, attrs_); END IF;
      Xml_Node (doc_, nd_nv_, 'cNvPr', 'xdr', attrs_);
      nd_cn_ := Xml_Node (doc_, nd_nv_, 'cNvPicPr', 'xdr');

      natr ('noChangeAspect', '1', attrs_);
      Xml_Node (doc_, nd_cn_, 'picLocks', 'a', attrs_);

      nd_bf_ := Xml_Node (doc_, nd_pi_, 'blipFill', 'xdr');

      natr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
      attr ('r:embed', rep ('rId:P1', to_char(drawing_.img_id)), attrs_);
      nd_bl_ := Xml_Node (doc_, nd_bf_, 'blip', 'a', attrs_);
      nd_et_ := Xml_Node (doc_, nd_bl_, 'extLst', 'a');

      natr ('uri', Get_Guid, attrs_);
      nd_el_ := Xml_Node (doc_, nd_et_, 'ext', 'a', attrs_);

      natr ('xmlns:a14', 'http://schemas.microsoft.com/office/drawing/2010/main', attrs_);
      attr ('val', '0', attrs_);
      Xml_Node (doc_, nd_el_, 'useLocalDpi', 'a14', attrs_);
      Xml_Node (doc_, Xml_Node(doc_,nd_bf_,'stretch','a'), 'fillRect', 'a');

      natr ('prst', 'rect', attrs_);
      Xml_Node (doc_, Xml_Node(doc_,nd_pi_,'spPr','xdr'), 'prstGeom', 'a', attrs_);
      Xml_Node (doc_, nd_tc_, 'clientData', 'xdr');

   END LOOP;

   Add1Xml (excel_, rep('xl/drawings/drawing:P1.xml',s_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

   << skip_drawings >>
   null;

END Finish_Ws_Drawings;

PROCEDURE Finish_Ws_Comments (
   excel_ IN OUT NOCOPY BLOB,
   s_     IN            PLS_INTEGER )
IS
   au_count_      PLS_INTEGER := 0;
   ws_authors_    tp_authors;
   author_        tp_author;
   doc_           dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_cms_        dbms_XmlDom.DomNode;
   nd_cml_        dbms_XmlDom.DomNode;
   nd_cm_         dbms_XmlDom.DomNode;
   nd_aus_        dbms_XmlDom.DomNode;
   nd_tx_         dbms_XmlDom.DomNode;
   nd_r_          dbms_XmlDom.DomNode;
   nd_pr_         dbms_XmlDom.DomNode;
   nd_xml_        dbms_XmlDom.DomNode;
   nd_sl_         dbms_XmlDom.DomNode;
   nd_st_         dbms_XmlDom.DomNode;
   nd_sh_         dbms_XmlDom.DomNode;
   nd_tb_         dbms_XmlDom.DomNode;
   nd_cd_         dbms_XmlDom.DomNode;
   attrs_         xml_attrs_arr;
   nl_            VARCHAR2(2);
   comment_w_rem_ NUMBER;
   comment_h_     NUMBER;
   col_w_         NUMBER;
   colspan_       NUMBER;
BEGIN

   IF wb_.sheets(s_).comments.count = 0 THEN
      goto skiop_comments;
   END IF;

   FOR c_ IN 1 .. wb_.sheets(s_).comments.count LOOP
      ws_authors_(wb_.sheets(s_).comments(c_).author) := 0;
   END LOOP;

   -- xl/comments:P1.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   nd_cms_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'comments', attrs_);
   nd_aus_ := Xml_Node (doc_, nd_cms_, 'authors');
   author_ := ws_authors_.first;
   WHILE author_ IS NOT null OR ws_authors_.next(author_) IS NOT null LOOP
      ws_authors_(author_) := au_count_;
      Xml_Text_Node (doc_, nd_aus_, 'author', author_);
      au_count_  := au_count_ + 1;
      author_ := ws_authors_.next(author_);
   END LOOP;

   nd_cml_ := Xml_Node (doc_, nd_cms_, 'commentList');
   FOR cm_ IN 1 .. wb_.sheets(s_).comments.count LOOP
      natr ('ref', Alfan_Cell (wb_.sheets(s_).comments(cm_).column, wb_.sheets(s_).comments(cm_).row), attrs_);
      attr ('authorId', ws_authors_(wb_.sheets(s_).comments(cm_).author), attrs_);
      nd_cm_ := Xml_Node (doc_, nd_cml_, 'comment', attrs_);
      nd_tx_ := Xml_Node (doc_, nd_cm_, 'text');
      IF wb_.sheets(s_).comments(cm_).author IS NOT null THEN
         nd_r_  := Xml_Node (doc_, nd_tx_, 'r');
         nd_pr_ := Xml_Node (doc_, nd_r_, 'rPr');
         Xml_Node (doc_, nd_pr_, 'b');

         natr ('val', '9', attrs_);
         Xml_Node (doc_, nd_pr_, 'sz', attrs_);

         natr ('indexed', '81', attrs_);
         Xml_Node (doc_, nd_pr_, 'color', attrs_);

         natr ('val', 'Tahoma', attrs_);
         Xml_Node (doc_, nd_pr_, 'rFont', attrs_);

         natr ('val', '1', attrs_);
         Xml_Node (doc_, nd_pr_, 'charset', attrs_);

         natr ('xml:space', 'preserve', attrs_);
         Xml_Text_Node (doc_, nd_r_, 't', wb_.sheets(s_).comments(cm_).author, attrs_);
      END IF;
      nd_r_  := Xml_Node (doc_, nd_tx_, 'r');
      nd_pr_ := Xml_Node (doc_, nd_r_, 'rPr');

      natr ('val', '9', attrs_);
      Xml_Node (doc_, nd_pr_, 'sz', attrs_);

      natr ('indexed', '81', attrs_);
      Xml_Node (doc_, nd_pr_, 'color', attrs_);

      natr ('val', 'Tahoma', attrs_);
      Xml_Node (doc_, nd_pr_, 'rFont', attrs_);

      natr ('val', '1', attrs_);
      Xml_Node (doc_, nd_pr_, 'charset', attrs_);

      natr ('xml:space', 'preserve', attrs_);
      nl_ := CASE WHEN wb_.sheets(s_).comments(cm_).author IS NOT null THEN chr(13) || chr(10) END;
      Xml_Text_Node (doc_, nd_r_, 't', nl_ || wb_.sheets(s_).comments(cm_).text, attrs_);
   END LOOP;

   Add1Xml (excel_, rep('xl/comments:P1.xml',s_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);


   -- xl/drawings/vmlDrawing:P1.vml
   doc_ := Dbms_XmlDom.newDomDocument;

   natr ('xmlns:v', 'urn:schemas-microsoft-com:vml', attrs_);
   attr ('xmlns:o', 'urn:schemas-microsoft-com:office:office', attrs_);
   attr ('xmlns:x', 'urn:schemas-microsoft-com:office:excel', attrs_);
   nd_xml_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'xml', attrs_);

   natr ('v:ext', 'edit', attrs_);
   nd_sl_ := Xml_Node (doc_, nd_xml_, 'shapelayout', 'o', attrs_);

   natr ('v:ext', 'edit', attrs_);
   attr ('data', '2', attrs_);
   Xml_Node (doc_, nd_sl_, 'idmap', 'o', attrs_);
   attrs_.delete;
   natr ('id', '_x0000_t202', attrs_);
   attr ('coordsize', '21600,21600', attrs_);
   attr ('o:spt', '202', attrs_);
   attr ('path', 'm,l,21600r21600,l21600,xe', attrs_);
   nd_st_ := Xml_Node (doc_, nd_xml_, 'shapetype', 'v', attrs_);

   natr ('joinstyle', 'miter', attrs_);
   Xml_Node (doc_, nd_st_, 'stroke', 'v', attrs_);

   natr ('gradientshapeok', 't', attrs_);
   attr ('o:connecttype', 'rect', attrs_);
   Xml_Node (doc_, nd_st_, 'path', 'v', attrs_);

   FOR cm_ IN 1 .. wb_.sheets(s_).comments.count LOOP

      natr ('id', rep('_x0000_s:P1', to_char(cm_)), attrs_);
      attr ('type', '#_x0000_t202', attrs_);
      attr ('style', rep ('position:absolute;margin-left:35.25pt;margin-top:3pt;z-index::P1;visibility:hidden;', to_char(cm_)), attrs_);
      attr ('fillcolor', '#ffffe1', attrs_);
      attr ('o:insetmode', 'auto', attrs_);
      nd_sh_ := Xml_Node (doc_, nd_xml_, 'shape', 'v', attrs_);

      natr ('color2', '#ffffe1', attrs_);
      Xml_Node (doc_, nd_sh_, 'fill', 'v', attrs_);

      natr ('n', 't', attrs_);
      attr ('color', 'black', attrs_);
      attr ('obscured', 't', attrs_);
      Xml_Node (doc_, nd_sh_, 'shadow', 'v', attrs_);

      natr ('o:connecttype', 'none', attrs_);
      Xml_Node (doc_, nd_sh_, 'path', 'v', attrs_);

      natr ('style', 'mso-direction-alt:auto', attrs_);
      nd_tb_ := Xml_Node (doc_, nd_sh_, 'textbox', 'v', attrs_);
      attr ('style', 'text-align:left', attrs_);
      Xml_Text_Node (doc_, nd_tb_, 'div', '', attrs_);

      natr ('ObjectType', 'Note', attrs_);
      nd_cd_ := Xml_Node (doc_, nd_sh_, 'ClientData', 'x', attrs_);
      Xml_Node (doc_, nd_cd_, 'MoveWithCells', 'x');
      Xml_Node (doc_, nd_cd_, 'SizeWithCells', 'x');

      comment_w_rem_ := wb_.sheets(s_).comments(cm_).width;
      comment_h_     := wb_.sheets(s_).comments(cm_).height;
      colspan_       := 1;
      LOOP
         IF wb_.sheets(s_).widths.exists(wb_.sheets(s_).comments(cm_).column+colspan_) THEN
            col_w_ := 256 * wb_.sheets(s_).widths(wb_.sheets(s_).comments(cm_).column+colspan_);
            col_w_ := trunc((col_w_+18)/256*7); -- assume default 11 point Calibri
         ELSE
            col_w_ := 64;
         END IF;
         EXIT WHEN comment_w_rem_ < col_w_;
         colspan_       := colspan_ + 1;
         comment_w_rem_ := comment_w_rem_ - col_w_;
      END LOOP;
      Xml_Text_Node (
         doc_, nd_cd_, 'Anchor',
         rep (
            ':P1,15,:P2,30,:P3,:P4,:P5,:P6',
            to_char(wb_.sheets(s_).comments(cm_).column),
            to_char(wb_.sheets(s_).comments(cm_).row),
            to_char(wb_.sheets(s_).comments(cm_).column+colspan_-1),
            to_char(round(comment_w_rem_)),
            to_char(wb_.sheets(s_).comments(cm_).row+1+trunc(comment_h_/20)),
            to_char(mod(comment_h_, 20))
         ), 'x'
      );
      Xml_Text_Node (doc_, nd_cd_, 'AutoFill', 'False', 'x');
      Xml_Text_Node (doc_, nd_cd_, 'Row', to_char(wb_.sheets(s_).comments(cm_).row-1), 'x');
      Xml_Text_Node (doc_, nd_cd_, 'Column', to_char(wb_.sheets(s_).comments(cm_).column-1), 'x');
   END LOOP;

   Add1Xml (excel_, rep('xl/drawings/vmlDrawing:P1.vml',s_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);


   << skiop_comments >>
   null;

end Finish_Ws_Comments;

FUNCTION Finish RETURN BLOB
IS
   excel_        BLOB;
   s_            PLS_INTEGER;
BEGIN

   -- Pad out the Pivot Cache before doing any Excel generation.  Pivot tables
   -- will need this data in a moment...
   Build_Pivot_Caches;

   Dbms_Lob.createTemporary (excel_, true);

   -- We need to sort out the Pivots first, because tables will inject data to
   -- sheets, which in turn will create some additional shared strings and the
   -- like.  All this needs to be modelled in PL/SQL data-structures before we
   -- go about building the XML for various parts
   Finish_Pivot_Caches (excel_);            -- xl/pivotCache/pivotCacheDefinition[1].xml, 
   Finish_Pivot_Tables (excel_);            -- xl/pivotTables/pivotTable[1].xml

   Finish_Content_Types (excel_);           -- [Content_Types].xml
   Finish_docProps (excel_);                -- docProps/core.xml
   Finish_Rels (excel_);                    -- _rels/.rels
   Finish_Shared_Strings (excel_);          -- xl/sharedStrings.xml
   Finish_Styles (excel_);                  -- xl/styles.xml
   Finish_Theme (excel_);                   -- xl/theme/theme1.xml
   Finish_Workbook (excel_);                -- xl/workbook.xml
   Finish_Workbook_Rels (excel_);           -- xl/_rels/workbook.xml.rels
   Finish_Drawings_Rels (excel_);           -- xl/drawings/_rels/drawing1.xml.rels

   s_ := wb_.sheets.first;
   WHILE s_ IS not null LOOP
      Finish_Worksheet (excel_, s_);        -- xl/worksheets/sheet:P1.xml
      Finish_Ws_Relationships (excel_, s_); -- xl/worksheets/_rels/sheet:P1.xml.rels
      Finish_Ws_Drawings (excel_, s_);      -- xl/drawings/drawing:P1.xml
      Finish_Ws_Comments (excel_, s_);      -- xl/drawings/vmlDrawing:P1.vml
      s_ := wb_.sheets.next(s_);
   END LOOP;

   Finish_Zip (excel_);
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
   widths_       tp_widths; --TYPE tp_widths IS TABLE OF NUMBER INDEX BY PLS_INTEGER;
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
                           Cell (col_, offset_+i_, value_dt_ => d_tab_(i_+d_tab_.first()), sheet_ => sh_);
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
                        Cell (col_, offset_+i_, value_str_ => v_tab_(i_+v_tab_.first()), sheet_ => sh_);
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
   cur_   INTEGER := Dbms_Sql.Open_Cursor;
   throw_ INTEGER;
BEGIN
   Dbms_Sql.Parse (cur_, sql_, dbms_sql.native);
   Do_Binding (cur_, binds_);
   throw_ := Dbms_Sql.Execute(cur_); -- ignore
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
   --gbp_curr_fmt_ VARCHAR2(200) := '_-£* #,##0_-;-£* #,##0_-;_-£* &quot;-&quot;_-;_-@_-';
   gbp_curr_fmt0_ VARCHAR2(200) := '_-"£"* #,##0_-;-"£"* #,##0_-;_-"£"* "-"_-;_-@_-';
   gbp_curr_fmt2_ VARCHAR2(200) := '_-"£"* #,##0.00_-;-"£"* #,##0.00_-;_-"£"* "-"_-;_-@_-';
BEGIN

   Clear_Workbook;
   New_Sheet ('Sheet 1');

   fonts_('head1')       := Get_Font (rgb_ => 'FFDBE5F1', bold_ => true);
   fonts_('bold')        := Get_Font (bold_ => true);
   fonts_('bold_lg')     := Get_Font (bold_ => true, fontsize_ => 14);
   fonts_('bld_wht')     := Get_Font (rgb_ => 'FFFFFFFF', bold_ => true);
   fonts_('bld_dk_bl')   := Get_Font (rgb_ => 'FF244062', bold_ => true);
   fonts_('bld_lt_bl')   := Get_Font (rgb_ => 'FFDCE6F1', bold_ => true);
   fonts_('bld_ltbl_lg') := Get_Font (rgb_ => 'FFDCE6F1', bold_ => true, fontsize_ => 14);
   fonts_('bld_lt_gr')   := Get_Font (rgb_ => 'FFEBF1DE', bold_ => true);
   fonts_('italic')      := Get_Font (italic_ => true);
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

   bdrs_('none')         := Get_Border ('none', 'none', 'none', 'none');
   bdrs_('dotted')       := Get_Border ('dotted', 'dotted', 'dotted', 'dotted');
   bdrs_('t_dotted')     := Get_Border ('dotted', 'none', 'none', 'none'); -- top, bottom, left, right
   bdrs_('tl_dotted')    := Get_Border ('dotted', 'none', 'dotted', 'none');
   bdrs_('tbl_dotted')   := Get_Border ('dotted', 'dotted', 'dotted', 'none');
   bdrs_('tr_dotted')    := Get_Border ('dotted', 'none', 'none', 'dotted');
   bdrs_('tb_dotted')    := Get_Border ('dotted', 'dotted', 'none', 'none');
   bdrs_('b_dotted')     := Get_Border ('none', 'dotted', 'none', 'none');
   bdrs_('bl_dotted')    := Get_Border ('none', 'dotted', 'dotted', 'none');
   bdrs_('l_dotted')     := Get_Border ('none', 'none', 'dotted', 'none');
   bdrs_('br_dotted')    := Get_Border ('none', 'dotted', 'none', 'dotted');
   bdrs_('r_dotted')     := Get_Border ('none', 'none', 'none', 'dotted');
   bdrs_('tbr_dotted')   := Get_Border ('dotted', 'dotted', 'none', 'dotted');
   bdrs_('thin')         := Get_Border ('thin', 'thin', 'thin', 'thin');
   bdrs_('t_thin')       := Get_Border ('thin', 'none', 'none', 'none'); -- top, bottom, left, right
   bdrs_('tl_thin')      := Get_Border ('thin', 'none', 'thin', 'none');
   bdrs_('tbl_thin')     := Get_Border ('thin', 'thin', 'thin', 'none');
   bdrs_('tr_thin')      := Get_Border ('thin', 'none', 'none', 'thin');
   bdrs_('tb_thin')      := Get_Border ('thin', 'thin', 'none', 'none');
   bdrs_('b_thin')       := Get_Border ('none', 'thin', 'none', 'none');
   bdrs_('bl_thin')      := Get_Border ('none', 'thin', 'thin', 'none');
   bdrs_('l_thin')       := Get_Border ('none', 'none', 'thin', 'none');
   bdrs_('br_thin')      := Get_Border ('none', 'thin', 'none', 'thin');
   bdrs_('r_thin')       := Get_Border ('none', 'none', 'none', 'thin');
   bdrs_('tbr_thin')     := Get_Border ('thin', 'thin', 'none', 'thin');
   bdrs_('medium')       := Get_Border ('medium', 'medium', 'medium', 'medium');
   bdrs_('t_medium')     := Get_Border ('medium', 'none', 'none', 'none'); -- top, bottom, left, right
   bdrs_('tl_medium')    := Get_Border ('medium', 'none', 'medium', 'none');
   bdrs_('tbl_medium')   := Get_Border ('medium', 'medium', 'medium', 'none');
   bdrs_('tr_medium')    := Get_Border ('medium', 'none', 'none', 'medium');
   bdrs_('tb_medium')    := Get_Border ('medium', 'medium', 'none', 'none');
   bdrs_('b_medium')     := Get_Border ('none', 'medium', 'none', 'none');
   bdrs_('bl_medium')    := Get_Border ('none', 'medium', 'medium', 'none');
   bdrs_('l_medium')     := Get_Border ('none', 'none', 'medium', 'none');
   bdrs_('br_medium')    := Get_Border ('none', 'medium', 'none', 'medium');
   bdrs_('r_medium')     := Get_Border ('none', 'none', 'none', 'medium');
   bdrs_('tbr_medium')   := Get_Border ('medium', 'medium', 'none', 'medium');
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
   numFmt_('0dp')        := Get_NumFmt ('#,##0');
   numFmt_('2dp')        := Get_NumFmt ('#,##0.00');
   numFmt_('dt_mid')     := Get_NumFmt ('dd mmm yyyy');
   numFmt_('dthm_mid')   := Get_NumFmt ('dd mmm yyyy hh:mm');
   numFmt_('dthma_mid')  := Get_NumFmt ('dd mmm yyyy hh:mm AM/PM');
   numFmt_('dthms_mid')  := Get_NumFmt ('dd mmm yyyy hh:mm:ss');
   numFmt_('dthmsa_mid') := Get_NumFmt ('dd mmm yyyy hh:mm:ss AM/PM');
   numFmt_('dt_long')    := Get_NumFmt ('dd mmmm yyyy');
   numFmt_('Mmm yyyy')   := Get_NumFmt ('Mmm yyyy');

   align_('left')        := Get_Alignment (vertical_ => 'center', horizontal_ => 'left',   wrapText_ => false);
   align_('leftw')       := Get_Alignment (vertical_ => 'center', horizontal_ => 'left',   wrapText_ => true);
   align_('right')       := Get_Alignment (vertical_ => 'center', horizontal_ => 'right',  wrapText_ => false);
   align_('center')      := Get_Alignment (vertical_ => 'center', horizontal_ => 'center', wrapText_ => false);
   align_('wrap')        := Get_Alignment (vertical_ => 'top',    horizontal_ => 'left',   wrapText_ => true);
   align_('wrap_r')      := Get_Alignment (vertical_ => 'top',    horizontal_ => 'right',  wrapText_ => true);

END Init_Workbook;

PROCEDURE Set_Param (
   params_ IN OUT params_arr,
   ix_     IN NUMBER,
   name_   IN VARCHAR2,
   val_    IN VARCHAR2,
   extra_  IN VARCHAR2 := '' )
IS BEGIN
   params_(ix_) := param_rec (
      param_name      => name_,
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
   sh_  PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN

   -- Information about the report is static, with the only option being as to
   -- whether we show the user who printed the report
   Cell (2, row_, 'Report Information', fontId_ => fonts_('head1'), fillId_ => fills_('dk_blue'), sheet_ => sh_);
   Cell (3, row_, '', fillId_ => fills_('dk_blue'), sheet_ => sh_);
   row_ := row_ + 1;
   Cell (2, row_, 'Report Name', fontId_ => fonts_('bold'), sheet_ => sh_);
   Cell (3, row_, value_str_ => report_name_);
   row_ := row_ + 1;
   Cell (2, row_, 'Executed at', fontId_ => fonts_('bold'), sheet_ => sh_);
   Cell (3, row_, value_str_ => to_char(sysdate, 'YYYY-MM-DD HH24:MI:SS'), sheet_ => sh_);
   row_ := row_ + 1;
   IF show_user_ THEN
      Cell (2, row_, 'Executed by', fontId_ => fonts_('bold'), sheet_ => sh_);
      Cell (3, row_, value_str_ => user, sheet_ => sh_);
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
      Cell (3, row_, value_str_ => params_(i_).param_value, sheet_ => sh_);
      Cell (4, row_, value_str_ => params_(i_).additional_info, sheet_ => sh_);
      row_ := row_ + 1;
   END LOOP;

   Set_Column_Width (2, 25, sh_);
   Set_Column_Width (3, 40, sh_);
   Set_Column_Width (4, 40, sh_);

END Create_Params_Sheet;

END as_xlsx;
/
