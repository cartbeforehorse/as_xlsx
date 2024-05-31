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

PROCEDURE Add_Fill (
   fill_id_     IN VARCHAR2,
   patternType_ IN VARCHAR2,
   fgRGB_       IN VARCHAR2 := null,
   bgRGB_       IN VARCHAR2 := null );

PROCEDURE Add_NumFmt (
   fmt_id_ IN VARCHAR2,
   format_ IN VARCHAR2 );

---------------------------------------
-- Alfan_Cell(), Alfan_Range()
--  Transforms a numeric cell or range reference into an Excel reference.  For
--  example [1, 2] becomes "A2"; [1, 2, 3, 8] becomes "A2:C8".  This is useful
--  when external code is trying to generate formulas.
--
FUNCTION Alfan_Cell (
   col_ IN PLS_INTEGER,
   row_ IN PLS_INTEGER ) RETURN VARCHAR2;

FUNCTION Alfan_Range (
   col_tl_ IN PLS_INTEGER,
   row_tl_ IN PLS_INTEGER,
   col_br_ IN PLS_INTEGER,
   row_br_ IN PLS_INTEGER ) RETURN VARCHAR2;


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
   sheet_         IN PLS_INTEGER := null );

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
   sheet_         IN PLS_INTEGER := null );

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


---------------------------------------
---------------------------------------
--
-- Type Definitions
--
--
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
TYPE tp_cell_value IS RECORD (
   str_val  VARCHAR2(32000),
   num_val  NUMBER,
   dt_val   DATE
);
TYPE tp_cell IS RECORD (
   datatype    VARCHAR2(30), -- string|number|date
   ora_value   tp_cell_value,
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

TYPE tp_pivot_info IS RECORD (
   on_page  PLS_INTEGER
   --osian
);
TYPE tp_pivots_dir IS TABLE OF tp_pivot_info index by PLS_INTEGER;
TYPE tp_pivot IS RECORD (
   pivot_name VARCHAR2(200)
);
TYPE tp_pivots IS TABLE OF tp_pivot index by PLS_INTEGER;

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
   fontid      PLS_INTEGER,
   pivots      tp_pivots,
   drawings    tp_drawings
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
   right  VARCHAR2(17)
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
TYPE tp_defined_names IS TABLE OF tp_defined_name index by PLS_INTEGER;

TYPE tp_image IS RECORD (
   img_blob    BLOB,
   img_hash    RAW(128), --NUMBER,
   width       PLS_INTEGER,
   height      PLS_INTEGER
);
TYPE tp_images IS TABLE OF tp_image index by PLS_INTEGER;

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
   fontid        PLS_INTEGER,
   pivots_list   tp_pivots_dir,
   images        tp_images
);

workbook              tp_book;
g_useXf_              BOOLEAN := true;
g_addtxt2utf8blob_tmp VARCHAR2(32767);

TYPE xml_attrs_arr IS TABLE OF VARCHAR2(2000) INDEX BY VARCHAR2(200);

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
   quiet_   IN BOOLEAN  := false )
IS BEGIN
   Cbh_Utils_API.Trace (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_, quiet_);
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

FUNCTION Make_Tag (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomElement
IS
   el_ dbms_XmlDom.DomElement;
   ix_ VARCHAR2(200) := attrs_.FIRST;
BEGIN
   el_ := CASE
      WHEN ns_ IS NOT null THEN Dbms_XmlDom.createElement (doc_, tag_name_, ns_)
      ELSE Dbms_XmlDom.createElement (doc_, tag_name_)
   END;
   WHILE ix_ IS NOT null LOOP
      Dbms_XmlDom.setAttribute (el_, ix_, attrs_(ix_));
      ix_ := attrs_.NEXT(ix_);
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
   name_        IN VARCHAR2,
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
---------------------------------------
--
-- Cell reference converters
-- > Alfanumeric to number reference.  Useful for generating formulas
--
--
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
   col_ IN PLS_INTEGER,
   row_ IN PLS_INTEGER ) RETURN VARCHAR2
IS
BEGIN
   RETURN Alfan_Col (col_) || to_char(row_);
END Alfan_Cell;

FUNCTION Alfan_Range (
   col_tl_ IN PLS_INTEGER,
   row_tl_ IN PLS_INTEGER,
   col_br_ IN PLS_INTEGER,
   row_br_ IN PLS_INTEGER ) RETURN VARCHAR2
IS BEGIN
   RETURN Alfan_Cell (col_tl_, row_tl_) || ':' || Alfan_Cell (col_br_, row_br_);
END Alfan_Range;

FUNCTION Col_Alfan(
   col_ IN VARCHAR2 ) RETURN PLS_INTEGER
IS BEGIN
   RETURN ascii(substr(col_,-1)) - 64
      + nvl((ascii(substr(col_,-2,1))-64) * 26, 0)
      + nvl((ascii(substr(col_,-3,1))-64) * 676, 0);
END Col_Alfan;


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
IS
   sh_ PLS_INTEGER  := nvl(sheet_, workbook.sheets.count());
BEGIN
   RETURN workbook.sheets(sh_).rows(row_)(col_).ora_value.num_val;
END Get_Cell_Value_Num;

FUNCTION Get_Cell_Value_Str (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null ) RETURN VARCHAR2
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   RETURN workbook.sheets(sh_).rows(row_)(col_).ora_value.str_val;
END Get_Cell_Value_Str;

FUNCTION Get_Cell_Value_Date (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null ) RETURN DATE
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   RETURN workbook.sheets(sh_).rows(row_)(col_).ora_value.dt_val;
END Get_Cell_Value_Date;

FUNCTION Get_Cell_Value (
   col_    IN PLS_INTEGER,
   row_    IN PLS_INTEGER,
   sheet_  IN PLS_INTEGER := null ) RETURN VARCHAR2
IS
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   IF workbook.sheets(sh_).rows(row_)(col_).datatype = 'string' THEN
      RETURN Get_Cell_Value_Str (col_, row_, sheet_);
   ELSIF workbook.sheets(sh_).rows(row_)(col_).datatype = 'number' THEN
      RETURN to_char(Get_Cell_VAlue_Num (col_, row_, sheet_));
   ELSIF workbook.sheets(sh_).rows(row_)(col_).datatype = 'date' THEN
      RETURN to_char (Get_Cell_Value_Date (col_, row_, sheet_), 'YYYY-MM-DD-HH24:MI');
   END IF;
END Get_Cell_Value;

---------------------------------------
---------------------------------------
--
-- Functions that build the internal PL/SQL model of the Excel sheet
--
--
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
      workbook.sheets(s_).drawings.delete();
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
   FOR i_ IN 1 .. workbook.images.count LOOP
      dbms_lob.freeTemporary (workbook.images(i_).img_blob);
   END LOOP;
   workbook.images.delete();
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
   throw_ PLS_INTEGER;
BEGIN
   throw_ := New_Sheet (sheetname_, tab_color_); --ignore
END New_Sheet;

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 )
IS BEGIN
   workbook.sheets(sheet_).name := nvl(dbms_xmlgen.convert(translate(name_, 'a/\[]*:?', 'a')), 'Sheet'||sheet_);
END Set_Sheet_Name;

PROCEDURE Set_Col_Width (
   sheet_  IN PLS_INTEGER,
   col_    IN PLS_INTEGER,
   format_ IN VARCHAR2 )
IS
   width_  NUMBER;
   nr_chr_ PLS_INTEGER;
BEGIN
   IF format_ IS null THEN
      return;
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
   sh_          PLS_INTEGER  := nvl(sheet_, workbook.sheets.count());
   Xf_          tp_Xf_fmt    := Get_Cell_Xff(sh_, col_, row_);
   cell_border_ tp_border    := workbook.borders(Xf_.borderId);
   cell_dt_     VARCHAR2(30) := workbook.sheets(sh_).rows(row_)(col_).datatype;
   border_id_   PLS_INTEGER;
BEGIN

   cell_border_.top    := nvl (top_,    cell_border_.top);
   cell_border_.bottom := nvl (bottom_, cell_border_.bottom);
   cell_border_.left   := nvl (left_,   cell_border_.left);
   cell_border_.right  := nvl (right_,  cell_border_.right);
   border_id_          := Get_Border (
      cell_border_.top, cell_border_.bottom, cell_border_.left, cell_border_.right
   );

   IF cell_dt_ = 'number' THEN
      Cell (
         col_, row_, Get_Cell_Value_Num (col_, row_, sh_), --workbook.sheets(sh_).rows(row_)(col_).ora_value.num_val,
         Xf_.numFmtId, Xf_.fontId, Xf_.fillId, border_id_, Xf_.alignment, sh_
      );
   ELSIF cell_dt_ = 'string' THEN
      Cell (
         col_, row_, Get_Cell_Value_Str (col_, row_, sh_), --workbook.sheets(sh_).rows(row_)(col_).ora_value.str_val,
         Xf_.numFmtId, Xf_.fontId, Xf_.fillId, border_id_, Xf_.alignment, sh_
      );
   ELSIF cell_dt_ = 'date' THEN
      Cell (
         col_, row_, Get_Cell_Value_Date (col_, row_, sh_), --workbook.sheets(sh_).rows(row_)(col_).ora_value.dt_val,
         Xf_.numFmtId, Xf_.fontId, Xf_.fillId, border_id_, Xf_.alignment, sh_
      );
   END IF;

END Add_Border_To_Cell;

-----
-- Add_Border_To_Range()
--   Take a range of cells and put a border around it!  The procedure will not
--   override other settings in that that range of cells even if some of those
--   other settings have set borders on some of the internal cells
--
PROCEDURE Add_Border_To_Range (
   cell_left_ IN PLS_INTEGER,
   cell_top_  IN PLS_INTEGER,
   width_     IN PLS_INTEGER,
   height_    IN PLS_INTEGER,
   style_     IN VARCHAR2    := 'medium', -- thin|medium|thick|dotted...
   sheet_     IN PLS_INTEGER := null )
IS
   sh_         PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
   col_start_  PLS_INTEGER := cell_left_;
   col_end_    PLS_INTEGER := cell_left_ + width_ - 1;
   row_start_  PLS_INTEGER := cell_top_;
   row_end_    PLS_INTEGER := cell_top_ + height_ - 1;
BEGIN

   -- for a 1 x 1 span...
   IF width_ = 1 AND height_ = 1 THEN
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
   Xf_ tp_Xf_fmt ) RETURN PLS_INTEGER
IS
   xf_count_ PLS_INTEGER := workbook.cellXfs.count();
   xfId_     PLS_INTEGER;
BEGIN
   FOR i_ IN 1 .. xf_count_ LOOP
      IF (   workbook.cellXfs(i_).numFmtId = Xf_.numFmtId
         AND workbook.cellXfs(i_).fontId = Xf_.fontId
         AND workbook.cellXfs(i_).fillId = Xf_.fillId
         AND workbook.cellXfs(i_).borderId = Xf_.borderId
         AND nvl(workbook.cellXfs(i_).alignment.vertical, 'x') = nvl (Xf_.alignment.vertical, 'x')
         AND nvl(workbook.cellXfs(i_).alignment.horizontal, 'x') = nvl (Xf_.alignment.horizontal, 'x')
         AND nvl(workbook.cellXfs(i_).alignment.wrapText, false) = nvl (Xf_.alignment.wrapText, false)
      ) THEN
         XfId_ := i_;
         exit;
      END IF;
   END LOOP;
   IF XfId_ IS null THEN -- we didn't find a matching style, so create a new one
      workbook.cellXfs(xf_count_+1) := Xf_;
      xfId_ := xf_count_ + 1;
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
   alignment_ IN tp_alignment := null ) RETURN VARCHAR2
IS
   XfId_   PLS_INTEGER;
   Xf_     tp_Xf_fmt;
   col_Xf_ tp_Xf_fmt;
   row_Xf_ tp_Xf_fmt;
BEGIN

   IF not g_useXf_ THEN
      RETURN '';
   END IF;

   IF workbook.sheets(sheet_).col_fmts.exists(col_) THEN
      col_Xf_ := workbook.sheets(sheet_).col_fmts(col_);
   END IF;
   IF workbook.sheets(sheet_).row_fmts.exists(row_) THEN
      row_Xf_ := workbook.sheets(sheet_).row_fmts(row_);
   END IF;

   Xf_.numFmtId  := coalesce (numFmtId_, col_Xf_.numFmtId, row_Xf_.numFmtId, workbook.sheets(sheet_).fontid, workbook.fontid);
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
      RETURN '';
   END IF;

   IF Xf_.numFmtId > 0 THEN
      Set_Col_Width (sheet_, col_, workbook.numFmts(workbook.numFmtIndexes(Xf_.numFmtId)).formatCode);
   END IF;

   XfId_ := Get_Or_Create_XfId (Xf_);
   RETURN 's="' || XfId_ || '"';

END Get_XfId;

FUNCTION Extract_Id_From_Style (
   style_ IN VARCHAR2 ) RETURN PLS_INTEGER
IS BEGIN
   RETURN CASE
      WHEN style_ IS null OR style_ = 't="s" ' THEN to_number(null)
      ELSE to_number(regexp_replace (style_, '.*s="(\d+)".*', '\1'))
   END;
END Extract_Id_From_Style;

FUNCTION Get_Cell_XfId (
   sheet_ IN PLS_INTEGER,
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER ) RETURN PLS_INTEGER
IS
   style_tag_ VARCHAR2(50);
BEGIN
   IF workbook.sheets(sheet_).rows.exists(row_) AND
      workbook.sheets(sheet_).rows(row_).exists(col_)
   THEN
      style_tag_ := workbook.sheets(sheet_).rows(row_)(col_).style;
   ELSE
      -- We need to create the cell in the PlSql model so that later functions
      -- can manipulate it
      CellB (col_, row_, sheet_ => sheet_);
   END IF;

   RETURN CASE
      WHEN style_tag_ IS null OR style_tag_ = 't="s"' THEN null
      ELSE Extract_Id_From_Style (style_tag_)
   END;
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
      RETURN workbook.cellXfs (xfId_);
   END IF;
END Get_Cell_Xf;

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
   -- If the Cell doesn't have its own style, then we also need to verify that
   -- the cell's Row + Column don't have a "background" style
   IF cell_XfId_ IS NOT null THEN

      RETURN workbook.cellXfs (cell_xfId_);

   ELSE

      IF workbook.sheets(sheet_).col_fmts.exists(col_) THEN
         col_Xf_ := workbook.sheets(sheet_).col_fmts(col_);
      END IF;
      IF workbook.sheets(sheet_).row_fmts.exists(row_) THEN
         row_Xf_ := workbook.sheets(sheet_).row_fmts(row_);
      END IF;

      Xf_.numFmtId  := coalesce (col_Xf_.numFmtId, row_Xf_.numFmtId, workbook.sheets(sheet_).fontid, workbook.fontid);
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
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).rows(row_)(col_).datatype  := 'number';
   workbook.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => '', num_val => value_, dt_val => null
   );
   workbook.sheets(sh_).rows(row_)(col_).value     := value_;
   workbook.sheets(sh_).rows(row_)(col_).style     := get_XfId (
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
   fm_ix_ PLS_INTEGER := workbook.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, workbook.sheets.count());
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
      workbook.formulas(fm_ix_) := formula_;
      workbook.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
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
   IF workbook.strings.exists(nvl(string_,'')) THEN
      ix_ := workbook.strings(nvl(string_,''));
   ELSE
      ix_ := workbook.strings.count();
      workbook.str_ind(ix_) := string_;
      workbook.strings(nvl(string_,'')) := ix_;
   END IF;
   workbook.str_cnt := workbook.str_cnt + 1;
   RETURN ix_;
END Add_String;

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
   sh_    PLS_INTEGER  := nvl(sheet_, workbook.sheets.count());
   align_ tp_alignment := alignment_;
BEGIN
   workbook.sheets(sh_).rows(row_)(col_).datatype  := 'string';
   workbook.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => value_, num_val => null, dt_val => null
   );
   workbook.sheets(sh_).rows(row_)(col_).value     := Add_String(value_);
   IF align_.wrapText IS null AND instr(value_, chr(13)) > 0 THEN
      align_.wrapText := true;
   END IF;
   workbook.sheets(sh_).rows(row_)(col_).style := 't="s" ' || get_XfId (
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
   fm_ix_ PLS_INTEGER := workbook.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, workbook.sheets.count());
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
      workbook.formulas(fm_ix_) := formula_;
      workbook.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
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
   sh_         PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).rows(row_)(col_).datatype  := 'date';
   workbook.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => '', num_val => null, dt_val => value_
   );
   workbook.sheets(sh_).rows(row_)(col_).value     := (value_ - date '1900-03-01') + 61;
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
   fm_ix_ PLS_INTEGER := workbook.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, workbook.sheets.count());
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
      workbook.formulas(fm_ix_) := formula_;
      workbook.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
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
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   Cell (col_, row_, value_, 0, sheet_ => sheet_);
   workbook.sheets(sh_).rows(row_)(col_).style := XfId_;
END Query_Date_Cell;

PROCEDURE Condition_Color_Col (
   col_   IN PLS_INTEGER,
   sheet_ IN PLS_INTEGER := null )
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

         XfId_ := Get_Cell_XfId (sh_, col_, r_);

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
   col_   IN PLS_INTEGER,
   row_   IN PLS_INTEGER,
   url_   IN VARCHAR2,
   value_ IN VARCHAR2    := null,
   sheet_ IN PLS_INTEGER := null )
IS
   ix_ PLS_INTEGER;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.sheets(sh_).rows(row_)(col_).value := add_string(nvl(value_, url_));
   workbook.sheets(sh_).rows(row_)(col_).style := 't="s" ' || get_XfId(sh_, col_, row_, '', Get_Font('Calibri', theme_ => 10, underline_ => true));
   ix_ := workbook.sheets(sh_).hyperlinks.count() + 1;
   workbook.sheets(sh_).hyperlinks(ix_).cell := alfan_col(col_) || row_;
   workbook.sheets(sh_).hyperlinks(ix_).url := url_;
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
   ix_ PLS_INTEGER := workbook.formulas.count;
   sh_ PLS_INTEGER := nvl (sheet_, workbook.sheets.count());
BEGIN
   workbook.formulas(ix_) := formula_;
   Cell (col_, row_, default_value_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sh_);
   workbook.sheets(sh_).rows(row_)(col_).formula_idx := ix_;
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
   ix_ PLS_INTEGER := workbook.formulas.count;
   sh_ PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
BEGIN
   workbook.formulas(ix_) := formula_;
   Cell (col_, row_, default_value_, numFmtId_, fontId_, fillId_, borderId_, alignment_, sh_);
   workbook.sheets(sh_).rows(row_)(col_).formula_idx := ix_;
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
      sheet_        => sheet_
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
   sh_         PLS_INTEGER := coalesce (sheet_, workbook.sheets.count);
   img_ix_     PLS_INTEGER;
   hash_       RAW(128) := Dbms_Crypto.Hash (img_blob_, dbms_crypto.hash_md5);
   img_rec_    tp_image;
   drawing_    tp_drawing;
   offset_     NUMBER;
   length_     NUMBER;
   file_chunk_ RAW(14);
   hex_        VARCHAR2(8);
BEGIN

   FOR i_ IN 1 .. workbook.images.count LOOP
      IF workbook.images(i_).img_hash = hash_ THEN
         img_ix_ := i_;
         exit;
      END IF;
   END LOOP;

   IF img_ix_ IS null THEN

      img_ix_ := workbook.images.count + 1;
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

      ELSIF utl_raw.substr (file_chunk_,1,2) = hextoraw('FFD8') -- SOI Start of Image
            AND rawtohex (utl_raw.substr(file_chunk_,3,2)) IN ('FFE0', 'FFE1') -- APP0 jpg; APP1 jpg
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

      ELSE -- unknown - use the values passed in
         img_rec_.width  := nvl(width_, 0);
         img_rec_.height := nvl(height_, 0);
      END IF;

      workbook.images(img_ix_) := img_rec_;

   END IF;

   drawing_.img_id      := img_ix_;
   drawing_.row         := row_;
   drawing_.col         := col_;
   drawing_.scale       := scale_;
   drawing_.name        := name_;
   drawing_.title       := title_;
   drawing_.description := description_;
   workbook.sheets(sh_).drawings(workbook.sheets(sh_).drawings.count+1) := drawing_;

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
   tl_col_     PLS_INTEGER, -- top left
   tl_row_     PLS_INTEGER,
   br_col_     PLS_INTEGER, -- bottom right
   br_row_     PLS_INTEGER,
   name_       VARCHAR2,
   sheet_      PLS_INTEGER := null,
   localsheet_ PLS_INTEGER := null )
IS
   ix_ PLS_INTEGER;
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


---------------------------------------
---------------------------------------
--
-- The Excel file's XML creators
--
--
PROCEDURE Finish_Content_Types (
   excel_ IN OUT NOCOPY BLOB )
IS
   s_        PLS_INTEGER;
   p_        PLS_INTEGER;
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_types_ dbms_XmlDom.DomNode;
   xml_type_ XmlType;
   attrs_    xml_attrs_arr;
BEGIN

   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   attrs_('xmlns') := 'http://schemas.openxmlformats.org/package/2006/content-types';
   nd_types_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Types', attrs_);

   attrs_.delete;
   attrs_('ContentType') := 'application/vnd.openxmlformats-package.relationships+xml';
   attrs_('Extension')   := 'rels';
   Xml_Node (doc_, nd_types_, 'Default', attrs_);
   attrs_('ContentType') := 'application/xml';
   attrs_('Extension')   := 'xml';
   Xml_Node (doc_, nd_types_, 'Default', attrs_);
   attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.vmlDrawing';
   attrs_('Extension')   := 'vml';
   Xml_Node (doc_, nd_types_, 'Default', attrs_);

   attrs_.delete;
   attrs_('PartName')    := '/xl/workbook.xml';
   attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
   Xml_Node (doc_, nd_types_, 'Override', attrs_);

   s_ := workbook.sheets.first;
   WHILE s_ IS NOT null LOOP
      attrs_('PartName')    := rep('/xl/worksheets/sheet:P1.xml', s_);
      attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
      Xml_Node (doc_, nd_types_, 'Override', attrs_);
      s_ := workbook.sheets.next(s_);
   END LOOP;

   attrs_('PartName')    := '/xl/theme/theme1.xml';
   attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.theme+xml';
   Xml_Node (doc_, nd_types_, 'Override', attrs_);
   attrs_('PartName')    := '/xl/styles.xml';
   attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
   Xml_Node (doc_, nd_types_, 'Override', attrs_);
   attrs_('PartName')    := '/xl/sharedStrings.xml';
   attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
   Xml_Node (doc_, nd_types_, 'Override', attrs_);

   p_ := workbook.pivots_list.first;
   WHILE p_ IS NOT null LOOP

      attrs_('PartName')    := rep('/xl/pivotCache/pivotCacheDefinition:P1.xml', p_);
      attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';
      Xml_Node (doc_, nd_types_, 'Override', attrs_);

      attrs_('PartName')    := rep('/xl/pivotCache/pivotCacheRecords:P1.xml', p_);
      attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml';
      Xml_Node (doc_, nd_types_, 'Override', attrs_);

      attrs_('PartName')    := rep('xl/pivotTables/pivotTable:P1.xml', p_);
      attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml';
      Xml_Node (doc_, nd_types_, 'Override', attrs_);

      p_ := workbook.pivots_list.next(p_);

   END LOOP;

   attrs_('PartName')    := '/docProps/core.xml';
   attrs_('ContentType') := 'application/vnd.openxmlformats-package.core-properties+xml';
   Xml_Node (doc_, nd_types_, 'Override', attrs_);
   attrs_('PartName')    := '/docProps/app.xml';
   attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
   Xml_Node (doc_, nd_types_, 'Override', attrs_);

   s_ := workbook.sheets.first;
   WHILE s_ IS NOT null LOOP
      IF workbook.sheets(s_).comments.count > 0 THEN
         attrs_('PartName')    := rep('/xl/comments:P1.xml', s_);
         attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
         Xml_Node (doc_, nd_types_, 'Override', attrs_);
      END IF;
      IF workbook.sheets(s_).drawings.count > 0 THEN
         attrs_('PartName')    := rep('/xl/drawings/drawing:P1.xml', s_);
         attrs_('ContentType') := 'application/vnd.openxmlformats-officedocument.drawing+xml';
         Xml_Node (doc_, nd_types_, 'Override', attrs_);
      END IF;
      s_ := workbook.sheets.next(s_);
   END LOOP;

   IF workbook.images.count > 0 THEN
      attrs_.delete;
      attrs_('ContentType') := 'image/png';
      attrs_('Extension')   := 'png';
      Xml_Node (doc_, nd_types_, 'Default', attrs_);
   END IF;

   xml_type_ := Dbms_XmlDom.getXmlType (doc_);
   Dbms_XmlDom.freeDocument (doc_);
   Add1Xml (excel_, '[Content_Types].xml', xml_type_.getClobVal);

END Finish_Content_Types;

PROCEDURE Finish_Rels (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_rels_  dbms_XmlDom.DomNode;
   attrs_    xml_attrs_arr;
BEGIN

   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   attrs_('xmlns') := 'http://schemas.openxmlformats.org/package/2006/relationships';
   nd_rels_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

   attrs_.delete;
   attrs_('Id')     := 'rId1';
   attrs_('Type')   := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
   attrs_('Target') := 'xl/workbook.xml';
   Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
   attrs_('Id')     := 'rId2';
   attrs_('Type')   := 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
   attrs_('Target') := 'docProps/core.xml';
   Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
   attrs_('Id')     := 'rId3';
   attrs_('Type')   := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
   attrs_('Target') := 'docProps/app.xml';
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
   attrs_('xmlns:cp')       := 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
   attrs_('xmlns:dc')       := 'http://purl.org/dc/elements/1.1/';
   attrs_('xmlns:dcterms')  := 'http://purl.org/dc/terms/';
   attrs_('xmlns:dcmitype') := 'http://purl.org/dc/dcmitype/';
   attrs_('xmlns:xsi')      := 'http://www.w3.org/2001/XMLSchema-instance';
   nd_cprop_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'coreProperties', 'cp', attrs_);

   Xml_Text_Node (doc_, nd_cprop_, 'creator',        sys_context('userenv','os_user'), 'dc');
   Xml_Text_Node (doc_, nd_cprop_, 'description',    rep('Build by version: :P1', VERSION_), 'dc');
   Xml_Text_Node (doc_, nd_cprop_, 'lastModifiedBy', sys_context('userenv','os_user'), 'cp');

   attrs_.delete;
   attrs_('xsi:type') := 'dcterms:W3CDTF';
   Xml_Text_Node (doc_, nd_cprop_, 'created',  to_char(current_timestamp,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM'), 'dcterms', attrs_);
   Xml_Text_Node (doc_, nd_cprop_, 'modified', to_char(current_timestamp,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM'), 'dcterms', attrs_);

   Add1Xml (excel_, 'docProps/core.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);


   -- docProps/app.xml
   doc_ := Dbms_XmlDom.newDomDocument;
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   attrs_.delete;
   attrs_('xmlns')    := 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
   attrs_('xmlns:vt') := 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
   nd_prop_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Properties', attrs_);

   Xml_Text_Node (doc_, nd_prop_, 'Application', 'Microsoft Excel');
   Xml_Text_Node (doc_, nd_prop_, 'DocSecurity', '0');
   Xml_Text_Node (doc_, nd_prop_, 'ScaleCrop', 'false');
   nd_hd_  := Xml_Node (doc_, nd_prop_, 'HeadingPairs');
   attrs_.delete;
   attrs_('size')     := '2';
   attrs_('baseType') := 'variant';
   nd_vec_ := Xml_Node (doc_, nd_hd_, 'vector', 'vt', attrs_);
   nd_var_ := Xml_Node (doc_, nd_vec_, 'variant', 'vt');
   Xml_Text_Node (doc_, nd_var_, 'lpstr', 'Worksheets', 'vt');
   nd_var_ := Xml_Node (doc_, nd_vec_, 'variant', 'vt');
   Xml_Text_Node (doc_, nd_var_, 'i4', to_char(workbook.sheets.count), 'vt');

   nd_top_ := Xml_Node (doc_, nd_prop_, 'TitlesOfParts');
   attrs_.delete;
   attrs_('size')     := workbook.sheets.count;
   attrs_('baseType') := 'lpstr';
   nd_vec_ := Xml_Node (doc_, nd_top_, 'vector', 'vt', attrs_);
   s_ := workbook.sheets.first;
   WHILE s_ IS NOT null LOOP
      Xml_Text_Node (doc_, nd_vec_, 'lpstr', workbook.sheets(s_).name, 'vt');
      s_ := workbook.sheets.next(s_);
   END LOOP;
   Xml_Text_Node (doc_, nd_prop_, 'LinksUpToDate', 'false');
   Xml_Text_Node (doc_, nd_prop_, 'SharedDoc', 'false');
   Xml_Text_Node (doc_, nd_prop_, 'HyperlinksChanged', 'false');
   Xml_Text_Node (doc_, nd_prop_, 'AppVersion', '14.0300');

   Add1Xml (excel_, 'docProps/app.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_docProps;

PROCEDURE Finish_Styles (
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
   xml_type_ XmlType;
   attrs_    xml_attrs_arr;
BEGIN
   -- xl/styles.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   attrs_('xmlns:cp')       := 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
   attrs_('xmlns:dc')       := 'http://purl.org/dc/elements/1.1/';
   attrs_('xmlns:dcterms')  := 'http://purl.org/dc/terms/';
   attrs_('xmlns:dcmitype') := 'http://purl.org/dc/dcmitype/';
   attrs_('xmlns:xsi')      := 'http://www.w3.org/2001/XMLSchema-instance';
   nd_cprop_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'coreProperties', 'cp', attrs_);
/*
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
  xxx_ := xxx_ || '</cellXfs>
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
</styleSheet>';
*/

   Add1Xml (excel_, 'xl/styles.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);
END Finish_Styles;

FUNCTION Finish_Drawing (
   drawing_ tp_drawing,
   ix_      PLS_INTEGER,
   sheet_   PLS_INTEGER ) RETURN VARCHAR2
IS
   rv_         VARCHAR2(32767);
   col_        PLS_INTEGER;
   row_        PLS_INTEGER;
   width_      NUMBER;
   height_     NUMBER;
   col_offs_   NUMBER;
   row_offs_   NUMBER;
   col_width_  NUMBER;
   row_height_ NUMBER;
   widths_     tp_widths;
   heights_    tp_row_fmts;
BEGIN
   width_  := workbook.images(drawing_.img_id).width;
   height_ := workbook.images(drawing_.img_id).height;
   IF drawing_.scale IS NOT null THEN
      width_  := drawing_.scale * width_;
      height_ := drawing_.scale * height_;
   END IF;
   IF workbook.sheets(sheet_).widths.count = 0 THEN
      -- assume default column widths!
      -- 64 px = 1 col = 609600
      col_ := trunc (width_/64);
      col_offs_ := (width_-col_*64)*9525;
      col_ := drawing_.col - 1 + col_;
   ELSE
      widths_ := workbook.sheets(sheet_).widths;
      col_ := drawing_.col;
      LOOP
         IF widths_.exists(col_) THEN
            col_width_ := round(7*widths_(col_));
         ELSE
            col_width_ := 64;
         END IF;
         EXIT WHEN width_ < col_width_;
         col_ := col_ + 1;
         width_ := width_ - col_width_;
      END LOOP;
      col_ := col_ - 1;
      col_offs_ := width_ * 9525;
   END IF;
   IF workbook.sheets(sheet_).row_fmts.count = 0 THEN
      -- assume default row heigths!
      -- 20 px = 1 row = 190500
      row_ := trunc (height_/20);
      row_offs_ := (height_- row_*20) * 9525;
      row_ := drawing_.row - 1 + row_;
   ELSE
      heights_ := workbook.sheets(sheet_).row_fmts;
      row_ := drawing_.row;
      LOOP
         IF heights_.exists(row_) AND heights_(row_).height IS NOT null THEN
            row_height_ := heights_(row_).height;
            row_height_ := round (4*row_height_/3);
         ELSE
            row_height_ := 20;
         END IF;
         EXIT WHEN height_ < row_height_;
         row_ := row_ + 1;
         height_ := height_ - row_height_;
      END LOOP;
      row_offs_ := height_ * 9525;
      row_ := row_ - 1;
   END IF;
   rv_ := '<xdr:twoCellAnchor editAs="oneCell">
<xdr:from><xdr:col>' || ( drawing_.col-1 ) || '</xdr:col>
<xdr:colOff>0</xdr:colOff>
<xdr:row>' || ( drawing_.row-1 ) || '</xdr:row>
<xdr:rowOff>0</xdr:rowOff>
</xdr:from>
<xdr:to>
<xdr:col>' || col_ || '</xdr:col>
<xdr:colOff>' || col_offs_ || '</xdr:colOff>
<xdr:row>' || row_ || '</xdr:row>
<xdr:rowOff>' || row_offs_ || '</xdr:rowOff>
</xdr:to>
<xdr:pic>
<xdr:nvPicPr>
<xdr:cNvPr id="3" name="' || coalesce (drawing_.name, 'Picture '||ix_) || '"';
   IF drawing_.title IS NOT null THEN
      rv_ := rv_ || ' title="' || drawing_.title || '"';
   END IF;
   IF drawing_.description IS NOT null THEN
      rv_ := rv_ || ' descr="' || drawing_.description || '"';
   END IF;
   rv_ := rv_ || '/>
<xdr:cNvPicPr>
<a:picLocks noChangeAspect="1"/>
</xdr:cNvPicPr>
</xdr:nvPicPr>
<xdr:blipFill>
<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' || drawing_.img_id || '">
<a:extLst>
<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
</a:ext>
</a:extLst>
</a:blip>
<a:stretch>
<a:fillRect/>
</a:stretch>
</xdr:blipFill>
<xdr:spPr>
<a:prstGeom prst="rect">
</a:prstGeom>
</xdr:spPr>
</xdr:pic>
<xdr:clientData/>
</xdr:twoCellAnchor>';
   RETURN rv_;
END Finish_Drawing;


PROCEDURE Finish_Ws_Relationships (
   excel_ IN OUT NOCOPY BLOB,
   s_     IN            PLS_INTEGER )
IS
   id_            PLS_INTEGER := 1;
   nr_hyperlinks_ PLS_INTEGER := workbook.sheets(s_).hyperlinks.count;
   nr_comments_   PLS_INTEGER := workbook.sheets(s_).comments.count;
   nr_pivots_     PLS_INTEGER := workbook.sheets(s_).pivots.count;
   nr_drawings_   PLS_INTEGER := workbook.sheets(s_).drawings.count;
   doc_           dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   xml_type_      XmlType;
   attrs_         xml_attrs_arr;
   nd_rels_       dbms_XmlDom.DomNode;
BEGIN
   IF nr_hyperlinks_ > 0 OR nr_comments_ > 0 OR nr_pivots_ > 0 OR nr_drawings_ > 0 THEN

      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
      attrs_('xmlns') := 'http://schemas.openxmlformats.org/package/2006/relationships';
      nd_rels_ := Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), 'Relationships', attrs_);

      IF nr_comments_ > 0 THEN
         attrs_.delete;
         attrs_('Id')     := rep('rId:P1', nr_hyperlinks_+2);
         attrs_('Type')   := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
         attrs_('Target') := rep ('../comments:P1.xml', s_);
         Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
         attrs_('Id')     := rep ('rId:P1', nr_hyperlinks_+1);
         attrs_('Type')   := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
         attrs_('Target') := rep ('../drawings/vmlDrawing:P1.vml', s_);
         Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
         id_ := id_ + 2;
      END IF;
      FOR h_ IN 1 .. nr_hyperlinks_ LOOP
         IF workbook.sheets(s_).hyperlinks(h_).url IS NOT null THEN
            attrs_('Id')         := rep ('rId:P1', h_);
            attrs_('Type')       := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
            attrs_('Target')     := workbook.sheets(s_).hyperlinks(h_).url;
            attrs_('TargetMode') := 'External';
            Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
            id_ := id_ + 1;
         END IF;
      END LOOP;
      IF nr_drawings_ > 0 THEN
         attrs_.delete;
         attrs_('Id')     := rep ('rId:P1', id_);
         attrs_('Type')   := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
         attrs_('Target') := rep ('../drawings/drawing:P1.xml', s_);
         Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      END IF;
      FOR p_ IN 1 .. nr_pivots_ LOOP
         attrs_.delete;
         attrs_('Id')     := rep ('rId:P1', id_);
         attrs_('Type')   := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition';
         attrs_('Target') := rep ('../pivotCache/pivotCacheDefinition:P1.xml',1);
         Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
         id_ := id_ + 1;
      END LOOP;

      xml_type_ := Dbms_XmlDom.getXmlType (doc_);
      Dbms_XmlDom.freeDocument (doc_);
      Add1Xml (excel_, rep('xl/worksheets/_rels/sheet:P1.xml.rels',s_), xml_type_.getClobVal);
   END IF;
END Finish_Ws_Relationships;


FUNCTION Finish RETURN BLOB
IS
   excel_        BLOB;
   yyy_          BLOB;
   xxx_          CLOB;
   formula_expr_ VARCHAR2(32767 char);
   c_            NUMBER;
   h_            NUMBER;
   w_            NUMBER;
   cw_           NUMBER;
   s_            PLS_INTEGER;
   row_ix_       PLS_INTEGER;
   col_ix_       PLS_INTEGER;
   col_min_      PLS_INTEGER;
   col_max_      PLS_INTEGER;

BEGIN
   Dbms_Lob.createTemporary (excel_, true);

   Finish_Content_Types (excel_);
   Finish_docProps (excel_);
   Finish_Rels (excel_);
   --Finish_Styles (excel_);

  -- xl/styles.xml
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
  xxx_ := xxx_ || '</cellXfs>
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
</styleSheet>';
  add1xml (excel_, 'xl/styles.xml', xxx_);

  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr date1904="false" defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
  s_ := workbook.sheets.first;
  WHILE s_ IS not null LOOP
    xxx_ := xxx_ || ('
<sheet name="' || workbook.sheets(s_).name || '" sheetId="' || s_ || '" r:id="rId' || ( 9 + s_ ) || '"/>' );
    s_ := workbook.sheets.next(s_);
  END LOOP;
  xxx_ := xxx_ || '</sheets>';
  IF workbook.defined_names.count() > 0 THEN
    xxx_ := xxx_ || '<definedNames>';
    FOR s_ IN 1 .. workbook.defined_names.count() LOOP
      xxx_ := xxx_ || ('
<definedName name="' || workbook.defined_names(s_).name || '"' ||
        CASE WHEN workbook.defined_names(s_).sheet IS not null THEN ' localSheetId="' || to_char(workbook.defined_names(s_).sheet) || '"' END ||
        '>' || workbook.defined_names(s_).ref || '</definedName>');
    END LOOP;
    xxx_ := xxx_ || '</definedNames>';
  END IF;
  xxx_ := xxx_ || '<calcPr calcId="144525"/></workbook>';
  add1xml (excel_, 'xl/workbook.xml', xxx_);

  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</a:theme>';
  add1xml (excel_, 'xl/theme/theme1.xml', xxx_);

  s_ := workbook.sheets.first;
  WHILE s_ IS not null LOOP
    col_min_ := 16384;
    col_max_ := 1;
    row_ix_ := workbook.sheets(s_).rows.first();
    WHILE row_ix_ IS not null LOOP
      col_min_ := least(col_min_, workbook.sheets(s_).rows(row_ix_).first());
      col_max_ := greatest(col_max_, workbook.sheets(s_).rows(row_ix_).last());
      row_ix_ := workbook.sheets(s_).rows.next(row_ix_);
    END LOOP;

    addtxt2utf8blob_init(yyy_);
    addtxt2utf8blob ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' ||
CASE WHEN workbook.sheets(s_).tabcolor IS not null THEN '<sheetPr><tabColor rgb="' || workbook.sheets(s_).tabcolor || '"/></sheetPr>' end ||
'<dimension ref="' || alfan_col(col_min_) || workbook.sheets(s_).rows.first() || ':' || alfan_col(col_max_) || workbook.sheets(s_).rows.last() || '"/>
<sheetViews>
<sheetView' || CASE WHEN s_ = 1 THEN ' tabSelected="1"' END || ' workbookViewId="0">', yyy_);
    IF workbook.sheets(s_).freeze_rows > 0 AND workbook.sheets(s_).freeze_cols > 0 THEN
      addtxt2utf8blob (
        '<pane xSplit="' || workbook.sheets(s_).freeze_cols || '" '
        || 'ySplit="' || workbook.sheets(s_).freeze_rows || '" '
        || 'topLeftCell="' || alfan_col(workbook.sheets(s_).freeze_cols+1) || (workbook.sheets(s_).freeze_rows+1) || '" '
        || 'activePane="bottomLeft" state="frozen"/>',
        yyy_
      );
    ELSE
      IF workbook.sheets(s_).freeze_rows > 0 THEN
        addtxt2utf8blob (
          '<pane ySplit="' || workbook.sheets(s_).freeze_rows || '" topLeftCell="A' ||
            (workbook.sheets(s_).freeze_rows+1) || '" activePane="bottomLeft" state="frozen"/>',
          yyy_
        );
      END IF;
      IF workbook.sheets(s_).freeze_cols > 0 THEN
        addtxt2utf8blob (
          '<pane xSplit="' || workbook.sheets(s_).freeze_cols || '" topLeftCell="' ||
          alfan_col(workbook.sheets(s_).freeze_cols+1) ||
          '1" activePane="bottomLeft" state="frozen"/>',
          yyy_
        );
      END IF;
    END IF;
    addtxt2utf8blob ('</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>', yyy_);
    IF workbook.sheets(s_).widths.count() > 0 THEN
      addtxt2utf8blob ('<cols>', yyy_);
      col_ix_ := workbook.sheets(s_).widths.first();
      WHILE col_ix_ IS not null LOOP
        addtxt2utf8blob ('<col min="' || col_ix_ || '" max="' || col_ix_ || '" width="' || to_char(workbook.sheets(s_).widths(col_ix_), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" customWidth="1"/>', yyy_);
        col_ix_ := workbook.sheets(s_).widths.next(col_ix_);
      END LOOP;
      addtxt2utf8blob('</cols>', yyy_);
    END IF;
    addtxt2utf8blob('<sheetData>', yyy_);
    row_ix_ := workbook.sheets(s_).rows.first();
    WHILE row_ix_ IS not null LOOP
      IF workbook.sheets(s_).row_fmts.exists(row_ix_) AND workbook.sheets(s_).row_fmts(row_ix_).height IS not null THEN
          addtxt2utf8blob( '<row r="' || row_ix_ || '" spans="' || col_min_ || ':' || col_max_ || '" customHeight="1" ht="'
                         || to_char( workbook.sheets(s_).row_fmts(row_ix_).height, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" >', yyy_ );
      ELSE
        addtxt2utf8blob( '<row r="' || row_ix_ || '" spans="' || col_min_ || ':' || col_max_ || '">', yyy_ );
      END IF;
      col_ix_ := workbook.sheets(s_).rows(row_ix_).first();
      WHILE col_ix_ IS not null LOOP
        IF workbook.sheets(s_).rows(row_ix_)(col_ix_).formula_idx IS null THEN
          formula_expr_ := null;
        ELSE
          formula_expr_ := '<f>' || workbook.formulas(workbook.sheets(s_).rows(row_ix_)(col_ix_).formula_idx) || '</f>';
        END IF;
        addtxt2utf8blob ('<c r="' || alfan_col(col_ix_) || row_ix_ || '"'
          || ' ' || workbook.sheets(s_).rows(row_ix_)(col_ix_).style
          || '>' || formula_expr_ || '<v>'
          || to_char(workbook.sheets(s_).rows(row_ix_)(col_ix_).value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )
          || '</v></c>', yyy_
        );
        col_ix_ := workbook.sheets(s_).rows(row_ix_).next(col_ix_);
      END LOOP;
      addtxt2utf8blob( '</row>', yyy_ );
      row_ix_ := workbook.sheets(s_).rows.next(row_ix_);
    END LOOP;
    addtxt2utf8blob( '</sheetData>', yyy_ );
    FOR a IN 1 ..  workbook.sheets(s_).autofilters.count() LOOP
      addtxt2utf8blob( '<autoFilter ref="' ||
        alfan_col( nvl( workbook.sheets(s_).autofilters(a).column_start, col_min_ ) ) ||
        nvl( workbook.sheets(s_).autofilters(a).row_start, workbook.sheets(s_).rows.first() ) || ':' ||
        alfan_col(coalesce( workbook.sheets(s_).autofilters(a).column_end, workbook.sheets(s_).autofilters(a).column_start, col_max_)) ||
        nvl(workbook.sheets(s_).autofilters(a).row_end, workbook.sheets(s_).rows.last()) || '"/>', yyy_);
    END LOOP;
    IF workbook.sheets(s_).mergecells.count() > 0 THEN
      addtxt2utf8blob( '<mergeCells count="' || to_char(workbook.sheets(s_).mergecells.count()) || '">', yyy_);
      FOR m IN 1 ..  workbook.sheets(s_).mergecells.count() LOOP
        addtxt2utf8blob( '<mergeCell ref="' || workbook.sheets(s_).mergecells( m ) || '"/>', yyy_);
      END LOOP;
      addtxt2utf8blob('</mergeCells>', yyy_);
    END IF;
--
    IF workbook.sheets(s_).validations.count() > 0 THEN
      addtxt2utf8blob( '<dataValidations count="' || to_char( workbook.sheets(s_).validations.count() ) || '">', yyy_ );
      FOR m IN 1 ..  workbook.sheets(s_).validations.count() LOOP
        addtxt2utf8blob ('<dataValidation' ||
            ' type="' || workbook.sheets(s_).validations(m).type || '"' ||
            ' errorStyle="' || workbook.sheets(s_).validations(m).errorstyle || '"' ||
            ' allowBlank="' || CASE WHEN nvl(workbook.sheets(s_).validations(m).allowBlank, true) THEN '1' ELSE '0' END || '"' ||
            ' sqref="' || workbook.sheets(s_).validations(m).sqref || '"', yyy_ );
        IF workbook.sheets(s_).validations(m).prompt IS not null THEN
          addtxt2utf8blob(' showInputMessage="1" prompt="' || workbook.sheets(s_).validations(m).prompt || '"', yyy_);
          IF workbook.sheets(s_).validations(m).title IS not null THEN
            addtxt2utf8blob( ' promptTitle="' || workbook.sheets(s_).validations(m).title || '"', yyy_);
          END IF;
        END IF;
        IF workbook.sheets(s_).validations(m).showerrormessage THEN
          addtxt2utf8blob( ' showErrorMessage="1"', yyy_);
          IF workbook.sheets(s_).validations(m).error_title IS not null THEN
            addtxt2utf8blob( ' errorTitle="' || workbook.sheets(s_).validations(m).error_title || '"', yyy_);
          END IF;
          IF workbook.sheets(s_).validations(m).error_txt IS not null THEN
            addtxt2utf8blob(' error="' || workbook.sheets(s_).validations(m).error_txt || '"', yyy_);
          END IF;
        END IF;
        addtxt2utf8blob( '>', yyy_ );
        IF workbook.sheets(s_).validations(m).formula1 IS not null THEN
          addtxt2utf8blob ('<formula1>' || workbook.sheets(s_).validations(m).formula1 || '</formula1>', yyy_);
        END IF;
        IF workbook.sheets(s_).validations(m).formula2 IS not null THEN
          addtxt2utf8blob ('<formula2>' || workbook.sheets(s_).validations(m).formula2 || '</formula2>', yyy_);
        END IF;
        addtxt2utf8blob ('</dataValidation>', yyy_);
      END LOOP;
      addtxt2utf8blob ('</dataValidations>', yyy_);
    END IF;

    IF workbook.sheets(s_).hyperlinks.count() > 0 THEN
      addtxt2utf8blob ('<hyperlinks>', yyy_);
      FOR h IN 1 ..  workbook.sheets(s_).hyperlinks.count() LOOP
        addtxt2utf8blob ('<hyperlink ref="' || workbook.sheets(s_).hyperlinks(h).cell || '" r:id="rId' || h || '"/>', yyy_);
      END LOOP;
      addtxt2utf8blob ('</hyperlinks>', yyy_);
    END IF;
    addtxt2utf8blob( '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', yyy_);
    IF workbook.sheets(s_).comments.count() > 0 THEN
      addtxt2utf8blob( '<legacyDrawing r:id="rId' || ( workbook.sheets(s_).hyperlinks.count() + 1 ) || '"/>', yyy_);
    END IF;
    IF workbook.sheets(s_).drawings.count > 0 THEN
      addtxt2utf8blob( '<drawing r:id="rId' || (workbook.sheets(s_).hyperlinks.count + sign(workbook.sheets(s_).comments.count)+1) || '"/>', yyy_);
    END IF;

    addtxt2utf8blob( '</worksheet>', yyy_);
    addtxt2utf8blob_finish(yyy_);
    add1file (excel_, rep('xl/worksheets/sheet:P1.xml',s_), yyy_);

    Finish_Ws_Relationships (excel_, s_);

    IF workbook.sheets(s_).drawings.count > 0 THEN
      xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
      FOR i_ IN 1 .. workbook.sheets(s_).drawings.count LOOP
        xxx_ := xxx_ || Finish_Drawing (workbook.sheets(s_).drawings(i_), i_, s_);
      END LOOP;
      xxx_ := xxx_ || '</xdr:wsDr>';
      add1xml (excel_, 'xl/drawings/drawing' || s_ || '.xml', xxx_);
    END IF;

    IF workbook.sheets(s_).comments.count() > 0 THEN
      DECLARE
        cnt PLS_INTEGER;
        author_ind tp_author;
      BEGIN
        authors.delete();
        FOR c IN 1 .. workbook.sheets(s_).comments.count() LOOP
          authors(workbook.sheets(s_).comments(c).author) := 0;
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
      FOR c IN 1 .. workbook.sheets(s_).comments.count() LOOP
        xxx_ := xxx_ || ( '<comment ref="' || alfan_col( workbook.sheets(s_).comments(c).column ) ||
           to_char (workbook.sheets(s_).comments(c).row || '" authorId="' || authors(workbook.sheets(s_).comments(c).author ) ) || '">
<text>');
        IF workbook.sheets(s_).comments(c).author IS not null THEN
          xxx_ := xxx_ || ( '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
             workbook.sheets(s_).comments(c).author || ':</t></r>' );
        END IF;
        xxx_ := xxx_ || ( '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
             CASE WHEN workbook.sheets(s_).comments(c).author IS not null THEN '
' END || workbook.sheets(s_).comments(c).text || '</t></r></text></comment>' );
      END LOOP;
      xxx_ := xxx_ || '</commentList></comments>';
      add1xml (excel_, 'xl/comments' || s_ || '.xml', xxx_);

      xxx_ := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
      FOR c IN 1 .. workbook.sheets(s_).comments.count() LOOP
        xxx_ := xxx_ || ( '<v:shape id="_x0000_s' || to_char(c) || '" type="#_x0000_t202"
style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' || to_char( c ) || ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' );
        w_ := workbook.sheets(s_).comments(c).width;
        c_ := 1;
        LOOP
          IF workbook.sheets(s_).widths.exists(workbook.sheets(s_).comments(c).column+c_) THEN
            cw_ := 256 * workbook.sheets(s_).widths(workbook.sheets(s_).comments(c).column+c_);
            cw_ := trunc((cw_+18)/256*7); -- assume default 11 point Calibri
          ELSE
            cw_ := 64;
          END IF;
          EXIT WHEN w_ < cw_;
          c_ := c_ + 1;
          w_ := w_ - cw_;
        END LOOP;
        h_ := workbook.sheets(s_).comments(c).height;
        xxx_ := xxx_ || ( '<x:Anchor>' || workbook.sheets(s_).comments(c).column || ',15,' ||
            workbook.sheets(s_).comments(c).row || ',30,' ||
            (workbook.sheets(s_).comments(c).column+c_-1) || ',' || round(w_) || ',' ||
            (workbook.sheets(s_).comments(c).row + 1 + trunc(h_/20)) || ',' || mod(h_, 20) || '</x:Anchor>' );
        xxx_ := xxx_ || ( '<x:AutoFill>False</x:AutoFill><x:Row>' ||
            (workbook.sheets(s_).comments(c).row-1) || '</x:Row><x:Column>' ||
            (workbook.sheets(s_).comments(c).column-1) || '</x:Column></x:ClientData></v:shape>' );
      END LOOP;
      xxx_ := xxx_ || '</xml>';
      add1xml (excel_, 'xl/drawings/vmlDrawing' || s_ || '.vml', xxx_);
    END IF;
--
    s_ := workbook.sheets.next(s_);
  END LOOP;

  xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
  s_ := workbook.sheets.first;
  WHILE s_ IS not null LOOP
    xxx_ := xxx_ || ( '
<Relationship Id="rId' || (9+s_) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || s_ || '.xml"/>' );
    s_ := workbook.sheets.next(s_);
  END LOOP;
  xxx_ := xxx_ || '</Relationships>';
  add1xml (excel_, 'xl/_rels/workbook.xml.rels', xxx_);

  IF workbook.images.count > 0 THEN
    xxx_ := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    FOR i_ IN 1 .. workbook.images.count LOOP
       add1file (excel_, 'xl/media/image' || i_ || '.png', workbook.images(i_).img_blob );
       xxx_ := xxx_ || ( '<Relationship Id="rId' || i_ || '" '
                    ||   'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image' || i_ || '.png'
                    ||   '"/>' );
    END LOOP;
    xxx_ := xxx_ || '</Relationships>';
    add1xml (excel_, 'xl/drawings/_rels/drawing1.xml.rels', xxx_);
  END IF;

  addtxt2utf8blob_init(yyy_);
  addtxt2utf8blob ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || workbook.str_cnt || '" uniqueCount="' || workbook.strings.count() || '">',
    yyy_
  );
  FOR i IN 0 .. workbook.str_ind.count() - 1 LOOP
    addtxt2utf8blob (
       '<si><t xml:space="preserve">' ||
       dbms_xmlgen.convert(substr(workbook.str_ind(i), 1, 32000)) || '</t></si>', yyy_
    );
  END LOOP;
  addtxt2utf8blob ('</sst>', yyy_);
  addtxt2utf8blob_finish(yyy_);
  add1file (excel_, 'xl/sharedStrings.xml', yyy_);
  finish_zip (excel_);
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
   gbp_curr_fmt0_ VARCHAR2(200) := '_-&#163;* #,##0_-;-&#163;* #,##0_-;_-&#163;* &quot;-&quot;_-;_-@_-';
   gbp_curr_fmt2_ VARCHAR2(200) := '_-&#163;* #,##0.00_-;-&#163;* #,##0.00_-;_-&#163;* &quot;-&quot;_-;_-@_-';
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
   sh_  PLS_INTEGER := nvl(sheet_, workbook.sheets.count());
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
