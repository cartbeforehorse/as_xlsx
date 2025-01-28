CREATE OR REPLACE PACKAGE Nyce_Xlsx IS

------------------------------------------------------------------------------
------------------------------------------------------------------------------
--
-- Author: Anton Scheffer
--   Website: http://technology.amis.nl/blog
--   See also: http://technology.amis.nl/blog/?p=10995
-- # License
--     Copyright (C) 2011 - 2024 by Anton Scheffer
--     See associated LICENSE.md file for details
--
-- Modifications added by Osian ap Garth since 2017, version-controlled since
-- 2021 in Git Hub:
--   >> https://github.com/cartbeforehorse/as_xlsx
-- Copyright(C) 2025 "Now you can, ey!" (nyce.software) as Nyce_Xlsx.
-- For usage notes and a bit of a discussion about the design changes between
-- Anton's version and this, see documentation in README.md
--
------------------------------------------------------------------------------
------------------------------------------------------------------------------


------------------------------------------------------------------------------
-- Coding Notes
--    We'd be grateful if you could keep to these principles:
--    - 3 space indents
--    - UPPERCASE Structural keywords (like FUNCTION, PROCEDURE, LOOP)
--    - UPPERCASE constant names
--    - Camel_Case_Underscored() function and procedure names, in order to let
--      them stand out a bit from variable names
--    - lowercase everything else
--    - variable names given trailing underscore_, and none of this p_, v_, i_
--      prefix silliness.  No need to name every variable a "variable" as it's
--      the least interesting property of any variable!!
--    - Commas go at the end of lines, not the start (as I am sure you do with
--      every other programming language in the world, including English)
--    - When calling functions with multi-line parameters, please use the left
--      convention from those two options below, not the right one!  It allows
--      better "at a glance" scanning of the code and indentation that doesn't
--      depend on the function name's length!
--
--        Package.Function (       Package.Function ( hello_   => 'hi',
--           hello_   => 'hi',                        bye_     => 'see ya',
--           bye_     => 'see ya',                    staying_ => 'for tea' );
--           staying_ => 'for tea
--        );
--
--    - Do not use code beautifiers; they make code ugly and seriously mess up
--      version-control
--    - Use AI by all means, but it's never really worked for me.  I prefer to
--      look up Excel structural questions on the MS Documentation pages which
--      are on these links.  Also, Excel conforms with ISO-29500-1:2016, and a
--      PDF document describing this standards can be downloaded here too:
--       => https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f780b2d6-8252-4074-9fe3-5d7bc4830968
--       => https://www.iso.org/standard/71691.html
--       => https://standards.iso.org/ittf/PubliclyAvailableStandards/index.html
--

--------------------------------------------------
-- Constants
--
DBMS_CRYPTO_INSTALLED_ CONSTANT BOOLEAN := true;

RANGE_DEFINED_NAME_ CONSTANT VARCHAR2(100) := 'DefinedName';
RANGE_TABLE_        CONSTANT VARCHAR2(100) := 'Table';

--------------------------------------------------
-- Public Types
--

-----
-- Excel pivot table axes must be outward facing so that external apps can use
-- them to define the correct structure.  We also publish types to represent a
-- cell's location, range and alignment.
--
TYPE tp_pivot_cols  IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;
TYPE tp_agg_fn IS RECORD (
   colid        PLS_INTEGER,
   agg_fn       VARCHAR2(20),   -- [sum,avg,count...]
   col_tot_name VARCHAR2(2000), -- 'Total' or 'Sum of xxx', depending on nr of aggregates
   col_agg_name VARCHAR2(2000)  -- 'Sum of xxx' where "xxx" is the col-name
);
TYPE tp_col_agg_fns IS TABLE OF tp_agg_fn INDEX BY PLS_INTEGER;  -- 1 based
TYPE tp_pivot_axes IS RECORD (
   hrollups    tp_pivot_cols,
   vrollups    tp_pivot_cols,
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
   range_type   VARCHAR2(11),   -- Table;DefinedName
   defined_name VARCHAR2(1000), -- 'MyDatacells'
   sheet_id     PLS_INTEGER,    -- sheet.name => My Perfect Sheet; nullable, a range doens't necessarily need a sheet
   tl           tp_cell_loc,    -- (2, 3, false, true)
   br           tp_cell_loc,    -- (6, 6, false, false) Alfan_Range() => 'My Perfect Sheet'!B$3:F6
   local_sheet  BOOLEAN,        -- DN only: sets the defined name to be accessible only on `sheet_id`
   style        VARCHAR2(1000), -- Tbl only: name of the table's style
   ws_rel       PLS_INTEGER,    -- Tbl only: relId in the worksheet rels file
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
TYPE params_arr IS TABLE OF param_rec INDEX BY PLS_INTEGER;


--------------------------------------------------
-- Fonts and fills stored by name.  By design these are globally accessibel to
-- the outside world
--
TYPE tp_fonts_list  IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(50);
TYPE tp_fills_list  IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(50);
TYPE tp_border_list IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(50);
TYPE tp_numFmt_list IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(50);
TYPE tp_align_list  IS TABLE OF tp_alignment INDEX BY VARCHAR2(50);
TYPE tp_xf_list     IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(50);
TYPE tp_numFmt_cols IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;

fonts_  tp_fonts_list;
fills_  tp_fills_list;
bdrs_   tp_border_list;
numFmt_ tp_numFmt_list;
align_  tp_align_list;
xf_     tp_xf_list;

--------------------------------------------------
-- Sheet setup functions and procedures
--
PROCEDURE Init_Workbook;

PROCEDURE Clear_Workbook;

FUNCTION New_Sheet (
   sheetname_      IN VARCHAR2    := null,
   tab_color_      IN VARCHAR2    := null,
   show_gridlines_ IN BOOLEAN     := null,
   grid_colour_ix_ IN PLS_INTEGER := null,
   show_headers_   IN BOOLEAN     := null ) RETURN PLS_INTEGER;

PROCEDURE New_Sheet (
   sheetname_      IN VARCHAR2    := null,
   tab_color_      IN VARCHAR2    := null,
   show_gridlines_ IN BOOLEAN     := null,
   grid_colour_ix_ IN PLS_INTEGER := null,
   show_headers_   IN BOOLEAN     := null );

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 );

PROCEDURE Set_Dft_Fmt_Date_Short (
   format_mask_ IN VARCHAR2 );
PROCEDURE Set_Dft_Fmt_Date_Long (
   format_mask_ IN VARCHAR2 );
PROCEDURE Set_Dft_Fmt_Date_Time (
   format_mask_ IN VARCHAR2 );
PROCEDURE Set_Dft_Fmt_Time (
   format_mask_ IN VARCHAR2 );
PROCEDURE Set_Dft_Fmt_Num (
   format_mask_ IN VARCHAR2 );
PROCEDURE Set_Dft_Fmt_Num_Dc (
   format_mask_ IN VARCHAR2 );

FUNCTION OraFmt2Excel (
   ora_fmt_in_ IN VARCHAR2 := null ) RETURN VARCHAR2;

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

FUNCTION Alfan_Range (
   range_ IN tp_cell_range ) RETURN VARCHAR2;

---------------------------------------
-- Get_Border()
--  Values allowed in all these parameters are as follows:
--    none;thin;medium;dashed;dotted;thick;double;hair;mediumDashed;
--    dashDot;mediumDashDot;dashDotDot;mediumDashDotDot;slantDashDot
--
FUNCTION Get_Border (
   top_        IN VARCHAR2 := 'thin',
   bottom_     IN VARCHAR2 := 'thin',
   left_       IN VARCHAR2 := 'thin',
   right_      IN VARCHAR2 := 'thin',
   rgb_top_    IN VARCHAR2 := '',
   rgb_bottom_ IN VARCHAR2 := '',
   rgb_left_   IN VARCHAR2 := '',
   rgb_right_  IN VARCHAR2 := '' ) RETURN PLS_INTEGER;
PROCEDURE Get_Border (
   top_        IN VARCHAR2 := 'thin',
   bottom_     IN VARCHAR2 := 'thin',
   left_       IN VARCHAR2 := 'thin',
   right_      IN VARCHAR2 := 'thin',
   rgb_top_    IN VARCHAR2 := '',
   rgb_bottom_ IN VARCHAR2 := '',
   rgb_left_   IN VARCHAR2 := '',
   rgb_right_  IN VARCHAR2 := '' );


PROCEDURE Add_Border_To_Range (
   col_start_ IN PLS_INTEGER,
   row_start_ IN PLS_INTEGER,
   col_end_   IN PLS_INTEGER,
   row_end_   IN PLS_INTEGER,
   style_     IN VARCHAR2    := 'medium', -- thin|medium|thick|dotted...
   rgb_       IN VARCHAR2    := '',
   sheet_     IN PLS_INTEGER := null );

-----
-- Get_Alignment()
--  Values allowed in vert/horiz: horizontal;center;centerContinuous;distributed;fill;general;justify;left;right
--  Values allowed in wrapText:   vertical;bottom;center;distributed;justify;top
--
FUNCTION Get_Alignment (
   vertical_   IN VARCHAR2 := null,
   horizontal_ IN VARCHAR2 := null,
   wrapText_   IN BOOLEAN  := null ) RETURN tp_alignment;

FUNCTION Get_XfId (
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null ) RETURN PLS_INTEGER;

PROCEDURE Cell ( -- NUMBER
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN NUMBER,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null );
PROCEDURE Cell (
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_num_  IN NUMBER,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null );
PROCEDURE CellN ( -- num version explicit
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_num_  IN NUMBER,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null );

PROCEDURE Cell ( -- VARCHAR
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN VARCHAR2,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null );
PROCEDURE Cell (
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_str_  IN VARCHAR2,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null );
PROCEDURE CellS ( -- string version overload
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_str_  IN VARCHAR2,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null );

PROCEDURE Cell ( -- DATE
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   value_     IN DATE,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null );
PROCEDURE Cell (
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_dt_   IN DATE,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null );
PROCEDURE CellD ( -- date version overload
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_dt_   IN DATE,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null );

PROCEDURE CellB ( -- empty
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   fillId_    IN PLS_INTEGER,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null );
PROCEDURE CellB ( -- empty overload
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null );


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

-----
-- Set_Table()
--  Values allowed in style_: TableStyleLight1;TableStyleLight21;TableStyleMedium1;TableStyleMedium28;TableStyleDark1;TableStyleDark11
--
PROCEDURE Set_Table (
   col_start_ PLS_INTEGER,
   col_end_   PLS_INTEGER,
   row_start_ PLS_INTEGER,
   row_end_   PLS_INTEGER,
   style_     VARCHAR2,
   tbl_name_  VARCHAR2    := null,
   sheet_     PLS_INTEGER := null );

PROCEDURE Set_Table (
   tbl_range_ tp_cell_range,
   style_     VARCHAR2,
   tbl_name_  VARCHAR2 := null );

PROCEDURE Set_Tabcolor (
   tabcolor_ VARCHAR2, -- hex Alpha-rgb value
   sheet_    PLS_INTEGER := null );

FUNCTION Finish (
   pw_ IN VARCHAR2 := '' ) RETURN BLOB;

PROCEDURE Save (
   directory_ IN VARCHAR2,
   filename_  IN VARCHAR2,
   pw_        IN VARCHAR2 := '' );

PROCEDURE Save (
   xl_blob_   IN BLOB,
   directory_ IN VARCHAR2,
   filename_  IN VARCHAR2 );

PROCEDURE Query2Sheet (
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_headers_ IN BOOLEAN        := true,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2Sheet (
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   col_headers_ IN BOOLEAN        := true,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2Sheet ( -- using REFCURSOR
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   rc_          IN OUT NOCOPY SYS_REFCURSOR,
   col_headers_ IN BOOLEAN        := true,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2SheetAndAutofilter ( -- with Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2SheetAndAutofilter ( -- no Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2SheetAndAutofilter ( -- ref-cursor
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   rc_          IN OUT SYS_REFCURSOR,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2Table ( -- with Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   table_style_ IN VARCHAR2,
   tbl_name_    IN VARCHAR2       := null,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2Table ( -- no Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   table_style_ IN VARCHAR2,
   tbl_name_    IN VARCHAR2       := null,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

PROCEDURE Query2Table ( -- ref-cursor
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   rc_          IN OUT SYS_REFCURSOR,
   table_style_ IN VARCHAR2,
   tbl_name_    IN VARCHAR2       := null,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() );

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
   extra_blurb_ IN VARCHAR2,
   show_user_   IN BOOLEAN     := true,
   sheet_       IN PLS_INTEGER := null );


END Nyce_Xlsx;
/
CREATE OR REPLACE PACKAGE BODY Nyce_Xlsx IS

VERSION_ CONSTANT VARCHAR2(20) := 'as_xlsx20';

LOCAL_FILE_HEADER_        CONSTANT RAW(4) := hextoraw('504B0304'); -- Local file header signature
END_OF_CENTRAL_DIRECTORY_ CONSTANT RAW(4) := hextoraw('504B0506'); -- End of central directory signature

CELL_DT_STRING_           CONSTANT VARCHAR2(10) := 'string';
CELL_DT_NUMBER_           CONSTANT VARCHAR2(10) := 'number';
CELL_DT_DATE_             CONSTANT VARCHAR2(10) := 'date';
CELL_DT_HYPERLINK_        CONSTANT VARCHAR2(10) := 'hyperlink';

-- These are default Excel formats, not Oracle!  These can get complicated and
-- long with specialised requirements, so allow for plenty of character space.
-- Each default has a corresponding "Set_()" procedure
dft_fmt_date_short_       VARCHAR2(200) := 'yyyy-mm-dd';
dft_fmt_date_long_        VARCHAR2(200) := 'Dy Mon yyyy';
dft_fmt_date_time_        VARCHAR2(200) := 'yyyy-mm-dd hh:mm';
dft_fmt_time_             VARCHAR2(200) := 'hh:mm'; -- "hh:mm:ss", "hh:mm AM/PM"
dft_fmt_num_              VARCHAR2(200) := '#,##0';
dft_fmt_num_dc_           VARCHAR2(200) := '#,##0.00';

---------------------------------------
---------------------------------------
--
-- Type Definitions
--
--

-----
-- formatting dtypes
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
   cell   VARCHAR2(10),
   url    VARCHAR2(1000),
   ws_rel PLS_INTEGER
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
TYPE tp_comments_list IS TABLE OF tp_comment INDEX BY PLS_INTEGER;
TYPE tp_comments IS RECORD (
   ws_rel        PLS_INTEGER,
   comments_list tp_comments_list
);

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
--   Sometimes we want to store data indexed by a unique string value.  But at
--   other times we need to keep that data ordered, which doesn't suit the way
--   that plsql-table-types work.  We therefore hold the data in two different
--   arrays, defined by these types.
--
TYPE tp_unique_data IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(32000);
TYPE tp_data_ix_ord IS TABLE OF VARCHAR2(32000) INDEX BY PLS_INTEGER;
TYPE tp_col_filters IS TABLE OF VARCHAR2(32000) INDEX BY PLS_INTEGER;
TYPE tp_col_cache_method IS TABLE OF VARCHAR2(20) INDEX BY PLS_INTEGER;

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
   flds_to_cache  tp_col_cache_method, -- dynamically built from `pivot_axes` on the Pivot Table
   cached_fields  tp_cache_fields,     -- dynamically built, indexed by col-heading
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
TYPE tp_tables_list  IS TABLE OF VARCHAR2(100) INDEX BY PLS_INTEGER;

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
TYPE tp_drawings_list IS TABLE OF tp_drawing INDEX BY PLS_INTEGER;
TYPE tp_drawings IS RECORD (
   ws_rel        PLS_INTEGER,
   drawings_list tp_drawings_list
);

-----
-- sheet type
--
TYPE tp_sheet IS RECORD (
   wb_rel         PLS_INTEGER,
   sheet_name     VARCHAR2(100),
   rows           tp_rows,
   widths         tp_widths,
   show_gridlines BOOLEAN,
   grid_colour_ix PLS_INTEGER,
   show_headers   BOOLEAN,
   tabcolor       VARCHAR2(8),
   fontId         PLS_INTEGER,
   freeze_rows    PLS_INTEGER,
   freeze_cols    PLS_INTEGER,
   autofilters    tp_autofilters,
   hyperlinks     tp_hyperlinks,
   col_fmts       tp_col_fmts,
   row_fmts       tp_row_fmts,
   comments       tp_comments,
   mergecells     tp_mergecells,
   validations    tp_validations,
   tables_list    tp_tables_list,
   pivots_list    tp_pivots_list,
   drawings       tp_drawings
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
   style  VARCHAR2(17),
   rgb    VARCHAR2(8)
);
TYPE tp_cell_borders IS RECORD (
   top    tp_border,
   bottom tp_border,
   left   tp_border,
   right  tp_border
);
TYPE tp_borders IS TABLE OF tp_cell_borders INDEX BY PLS_INTEGER;
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
   fontId        PLS_INTEGER,
   fills         tp_fills,
   borders       tp_borders,
   numFmts       tp_numFmts,
   cellXfs       tp_cellXfs,
   formulas      tp_formulas,
   defined_names tp_defined_names, -- defined-range-names + tables
   tables_list   tp_tables_list,   -- [1 => 'Table1'], referencing defined_names
   pivot_caches  tp_pivot_caches,
   pivot_tables  tp_pivot_tables,
   images        tp_images
);


wb_                   tp_book;
g_addtxt2utf8blob_tmp VARCHAR2(32767);


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
   Nyce_Utils.Raise_App_Error (
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
   Nyce_Utils.Trace (m_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_, quiet_);
END Trace;
PROCEDURE Debug (
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
   indent_  IN NUMBER   := 0 )
IS BEGIN
   Dbms_Output.Put_Line ('--- debugging...');
   Trace (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_, true, indent_);
END Debug;
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
   RETURN Nyce_Utils.Rep (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_);
END Rep;


---------------------------------------
---------------------------------------
--
-- Excel helpers
--
--
PROCEDURE Set_Dft_Fmt_Date_Short (
   format_mask_ IN VARCHAR2 )
IS BEGIN
   dft_fmt_date_short_ := format_mask_;
END Set_Dft_Fmt_Date_Short;

PROCEDURE Set_Dft_Fmt_Date_Long (
   format_mask_ IN VARCHAR2 )
IS BEGIN
   dft_fmt_date_long_ := format_mask_;
END Set_Dft_Fmt_Date_Long;

PROCEDURE Set_Dft_Fmt_Date_Time (
   format_mask_ IN VARCHAR2 )
IS BEGIN
   dft_fmt_date_time_ := format_mask_;
END Set_Dft_Fmt_Date_Time;

PROCEDURE Set_Dft_Fmt_Time (
   format_mask_ IN VARCHAR2 )
IS BEGIN
   dft_fmt_time_ := format_mask_;
END Set_Dft_Fmt_Time;

PROCEDURE Set_Dft_Fmt_Num (
   format_mask_ IN VARCHAR2 )
IS BEGIN
   dft_fmt_num_ := format_mask_;
END Set_Dft_Fmt_Num;

PROCEDURE Set_Dft_Fmt_Num_Dc (
   format_mask_ IN VARCHAR2 )
IS BEGIN
   dft_fmt_num_dc_ := format_mask_;
END Set_Dft_Fmt_Num_Dc;

FUNCTION Get_Guid RETURN VARCHAR2
IS
   guid_ VARCHAR2(50) := RawToHex(sys_guid());
BEGIN
   RETURN '{' ||
      substr (guid_, 1,  8) || '-' || substr (guid_, 9,  4) || '-' ||
      substr (guid_, 13, 4) || '-' || substr (guid_, 17, 4) || '-' ||
      substr (guid_, 21, 12) || '}';
END Get_Guid;

-----
-- Name_Checker()
--   In Excel Formula tab you'll find the defined names section.  Each defined
--   name must comply with Excel's variable-naming convention, and should also
--   be unique within its scope.  This function checks those requirements.
FUNCTION Name_Checker (
   proposed_name_ IN VARCHAR2,
   name_type_     IN VARCHAR2 := RANGE_DEFINED_NAME_ ) RETURN VARCHAR2
IS
   ret_name_ VARCHAR2(2000) := proposed_name_;
BEGIN
   IF proposed_name_ IS NOT null THEN
      IF not regexp_like (proposed_name_, '^[a-zA-Z_]') THEN
         Raise_App_Error ('A registered name must start with a letter or an underscore.');
      ELSIF not regexp_like (proposed_name_, '^[a-zA-Z0-9\._\]+$') THEN
         Raise_App_Error (
            'A registered name must not contain spaces or any operation characters, such as ' ||
            'plus (+), divide (/) etc.  To keep it simple, use alpha-numeric and underscore only!'
         );
      END IF;
   ELSIF proposed_name_ IS null AND name_type_ = RANGE_DEFINED_NAME_ THEN
      Raise_App_Error ('A defined name may not have an empty descriptor!');
   ELSIF proposed_name_ IS null AND name_type_ = RANGE_TABLE_ THEN
      ret_name_ := 'Table' || to_char(wb_.tables_list.count+1);
   END IF;
   
   IF wb_.defined_names.exists(ret_name_) THEN
      Raise_App_Error ('Defined name ":P1" already exists on this workbook.', ret_name_);
   END IF;
   -- This duplication check is still a little naive.  Defined Names can exist
   -- in different scopes, meaning they either exist globally on the workbook,
   -- or locally to a specific sheet.  For now, this function allows us to use
   -- each DN only once on the entire workbook, but we will need to look again
   -- at how this works at a later date.
   RETURN ret_name_;
END Name_Checker;


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
   Dbms_Lob.createTemporary (blob_, true);
END addtxt2utf8blob_init;

PROCEDURE Addtxt2utf8blob_Finish (
   blob_ IN OUT NOCOPY BLOB )
IS
   raw_ RAW(32767);
BEGIN
   raw_ := utl_i18n.string_to_raw (g_addtxt2utf8blob_tmp, 'AL32UTF8');
   Dbms_Lob.writeAppend (blob_, utl_raw.length(raw_), raw_);
EXCEPTION
   WHEN value_error THEN
      raw_ := utl_i18n.string_to_raw(substr(g_addtxt2utf8blob_tmp,1,16381), 'AL32UTF8');
      Dbms_Lob.writeAppend (blob_, utl_raw.length(raw_), raw_);
      raw_ := utl_i18n.string_to_raw(substr(g_addtxt2utf8blob_tmp,16382), 'AL32UTF8');
      Dbms_Lob.writeAppend (blob_, utl_raw.length(raw_), raw_);
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
   FOR i_ IN 0 .. trunc((Dbms_Lob.getLength(blob_)-1)/len_) LOOP
      Utl_File.Put_Raw (fh_, Dbms_Lob.Substr(blob_, len_, i_*len_+1));
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
   IF big_ < 0 THEN
      RETURN Utl_Raw.Reverse (to_char(4294967296+big_, 'fm0XXXXXXX'));
   ELSE
      RETURN Utl_Raw.Reverse (to_char(big_, substr('fm0XXXXXXXXXXXXXXXXXXX', 1, 2+(2*bytes_))));
   END IF;
END Little_Endian;

FUNCTION Little_Endian (
   num_   RAW,
   pos_   PLS_INTEGER := 1,
   bytes_ PLS_INTEGER := null ) RETURN INTEGER
IS BEGIN
   RETURN to_number (
      Utl_Raw.Reverse (Utl_Raw.Substr (num_, pos_, bytes_)),
      'XXXXXXXXXXXXXXXX'
   );
END Little_Endian;

FUNCTION Blob2Num (
   blob_ BLOB,
   len_  INTEGER,
   pos_  INTEGER ) RETURN NUMBER
IS BEGIN
   RETURN utl_raw.cast_to_binary_integer (
      Dbms_Lob.Substr (blob_, len_, pos_), utl_raw.little_endian
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
   len_ := nvl(Dbms_Lob.GetLength(content_), 0);
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
      Dbms_Lob.createTemporary (zipped_blob_, true);
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
      Dbms_Lob.Copy (zipped_blob_, blob_, clen_, Dbms_Lob.getLength(zipped_blob_)+1, 11); -- compressed content
   ELSIF clen_ > 0 THEN
      Dbms_Lob.Copy (zipped_blob_, blob_, clen_, Dbms_Lob.getLength(zipped_blob_)+1, 1); --  content
   END IF;
   IF Dbms_Lob.isTemporary(blob_) = 1 THEN
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
   offs_dir_header_ := Dbms_Lob.getLength (zipped_blob_);
   offset_ := 1;
   WHILE Dbms_Lob.Substr(zipped_blob_, Utl_Raw.Length(LOCAL_FILE_HEADER_), offset_) = LOCAL_FILE_HEADER_ LOOP
      nr_ := nr_ + 1;
      Dbms_Lob.Append (
         zipped_blob_,
         Utl_Raw.Concat (
            hextoraw('504B0102'),      -- Central directory file header signature
            hextoraw('1400'),          -- version 2.0
            Dbms_Lob.Substr(zipped_blob_, 26, offset_+4),
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
            Dbms_Lob.Substr(zipped_blob_, blob2num(zipped_blob_,2,offset_+26),offset_+30) -- File name
         )
      );
      offset_ := offset_ + 30 +
         blob2num (zipped_blob_, 4, offset_+18 ) + -- compressed size
         blob2num (zipped_blob_, 2, offset_+26 ) + -- File name length
         blob2num (zipped_blob_, 2, offset_+28 );  -- Extra field length
   END LOOP;
   offs_end_header_ := Dbms_Lob.getLength(zipped_blob_);
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
-- > Alfan_Col()   => helps to convert (2, 3) => B3; actually ports AA => 27
-- > Col_Alfan()   => helps to convert B3 => (2, 3); actually ports 27 => AA
-- > Alfan_Cell()  => (2, 3, true, false) => B$3
-- > Alfan_Range() => (2, 3, 7, 7) => B3:G7
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
   range_ IN tp_cell_range ) RETURN VARCHAR2
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
      wb_.sheets(sheet_).sheet_name, col_tl_, row_tl_, col_br_, row_br_,
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
   RETURN wb_.sheets(sheet_).sheet_name;
END Sheet_Name;

FUNCTION Sheet_Name (
   range_ IN tp_cell_range ) RETURN VARCHAR2
IS BEGIN
   RETURN wb_.sheets(range_.sheet_id).sheet_name;
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
   range_     IN OUT NOCOPY tp_cell_range,
   sheet_     IN PLS_INTEGER := null,
   allow_dup_ IN BOOLEAN     := true )
IS
   i_       PLS_INTEGER := 1;
   row_     PLS_INTEGER := range_.tl.r;
   sh_      PLS_INTEGER := coalesce (range_.sheet_id, sheet_, wb_.sheets.count);
   new_val_ VARCHAR2(32000);
   uq_      tp_unique_data;
BEGIN
   IF range_.col_names.count = 0 THEN -- check that `col_names` is not already filled out
      FOR c_ IN range_.tl.c .. range_.br.c LOOP
         new_val_ := wb_.sheets(sh_).rows(row_)(c_).ora_value.str_val;
         IF not allow_dup_ THEN
            IF uq_.exists(new_val_) THEN
               Raise_App_Error ('Heading :P1 may only appear once in this range!', new_val_);
            END IF;
            uq_(new_val_) := 1;
         END IF;
         range_.col_names(i_) := new_val_;
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

PROCEDURE Build_Si_From_Range (
   unq_data_ IN OUT NOCOPY tp_unique_data,
   ord_data_ IN OUT NOCOPY tp_data_ix_ord,
   range_    IN tp_cell_range,
   col_offs_ IN PLS_INTEGER,
   sheet_    IN PLS_INTEGER := null )
IS
   col_       PLS_INTEGER := range_.tl.c + col_offs_ - 1;
   row_start_ PLS_INTEGER := range_.tl.r + 1; -- allow for header row
   row_end_   PLS_INTEGER := range_.br.r;
   sh_        PLS_INTEGER := coalesce (range_.sheet_id, sheet_, wb_.sheets.count);
   val_       VARCHAR2(2000);
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
END Build_Si_From_Range;


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
         wb_.sheets(s_).rows(row_ix_).delete;
         row_ix_ := wb_.sheets(s_).rows.next(row_ix_);
      END LOOP;
      wb_.sheets(s_).rows.delete;
      wb_.sheets(s_).widths.delete;
      wb_.sheets(s_).autofilters.delete;
      wb_.sheets(s_).hyperlinks.delete;
      wb_.sheets(s_).col_fmts.delete;
      wb_.sheets(s_).row_fmts.delete;
      wb_.sheets(s_).comments.comments_list.delete;
      wb_.sheets(s_).comments := null;
      wb_.sheets(s_).mergecells.delete;
      wb_.sheets(s_).validations.delete;
      wb_.sheets(s_).tables_list.delete;
      wb_.sheets(s_).pivots_list.delete;
      wb_.sheets(s_).drawings.drawings_list.delete;
      wb_.sheets(s_).drawings := tp_drawings();
      s_ := wb_.sheets.next(s_);
   END LOOP;
   wb_.strings.delete;
   wb_.str_ind.delete;
   wb_.fonts.delete;
   wb_.fills.delete;
   wb_.borders.delete;
   wb_.numFmts.delete;
   wb_.cellXfs.delete;
   wb_.formulas.delete;
   wb_.defined_names.delete;
   wb_.tables_list.delete;
   FOR i_ IN 1 .. wb_.images.count LOOP
      Dbms_Lob.freeTemporary (wb_.images(i_).img_blob);
   END LOOP;
   wb_.images.delete;
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
   sheetname_      IN VARCHAR2    := null,
   tab_color_      IN VARCHAR2    := null,
   show_gridlines_ IN BOOLEAN     := null,
   grid_colour_ix_ IN PLS_INTEGER := null, -- index in default color palette 0 - 55
   show_headers_   IN BOOLEAN     := null ) RETURN PLS_INTEGER
IS
   s_          PLS_INTEGER   := wb_.sheets.count + 1;
   sheet_name_ VARCHAR2(100) := nvl (sheetname_, 'Sheet ' || s_);
BEGIN
   wb_.sheets(s_).sheet_name := nvl (
      Dbms_XmlGen.Convert(translate(sheet_name_, 'a/\[]*:?', 'a')),
      'Sheet' || s_
   );
   wb_.sheets(s_).show_gridlines := show_gridlines_;
   wb_.sheets(s_).grid_colour_ix := grid_colour_ix_;
   wb_.sheets(s_).show_headers   := show_headers_;
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
   sheetname_      IN VARCHAR2    := null,
   tab_color_      IN VARCHAR2    := null,
   show_gridlines_ IN BOOLEAN     := null,
   grid_colour_ix_ IN PLS_INTEGER := null,
   show_headers_   IN BOOLEAN     := null )
IS
   throw_ PLS_INTEGER;
BEGIN
   throw_ := New_Sheet (
      sheetname_, tab_color_, show_gridlines_, grid_colour_ix_, show_headers_
   );
END New_Sheet;

PROCEDURE Set_Sheet_Name (
   sheet_  IN PLS_INTEGER,
   name_   IN VARCHAR2 )
IS BEGIN
   wb_.sheets(sheet_).sheet_name := nvl (
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

-----
-- OraFmt2Excel()
--  Changes date and number formats from Oracle to Excel.  Clever, but I'm not
--  quite sure what the use-case is.
FUNCTION OraFmt2Excel (
   ora_fmt_in_ VARCHAR2 := null ) RETURN VARCHAR2
IS
   ora_fmt_ VARCHAR2(1000) := substr (ora_fmt_in_, 1, 1000);
BEGIN
   ora_fmt_ := replace(replace(ora_fmt_,'hh24','hh'),'hh12','hh');
   ora_fmt_ := replace( ora_fmt_, 'mi', 'mm' );
   ora_fmt_ := replace( replace( replace( ora_fmt_, 'AM', '~~' ), 'PM', '~~' ), '~~', 'AM/PM' );
   ora_fmt_ := replace( replace( replace( ora_fmt_, 'am', '~~' ), 'pm', '~~' ), '~~', 'AM/PM' );
   ora_fmt_ := replace( replace( ora_fmt_, 'day', 'DAY' ), 'DAY', 'dddd' );
   ora_fmt_ := replace( replace( ora_fmt_, 'dy', 'DY' ), 'DAY', 'ddd' );
   ora_fmt_ := replace( replace( ora_fmt_, 'RR', 'RR' ), 'RR', 'YY' );
   ora_fmt_ := replace( replace( ora_fmt_, 'month', 'MONTH' ), 'MONTH', 'mmmm' );
   ora_fmt_ := replace( replace( ora_fmt_, 'mon', 'MON' ), 'MON', 'mmm' );
   ora_fmt_ := replace( ora_fmt_, '9', '#' );
   ora_fmt_ := replace( ora_fmt_, 'D', '.' );
   ora_fmt_ := replace( ora_fmt_, 'G', ',' );
   RETURN ora_fmt_;
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
   throw_ := Get_Fill (patternType_, fgRGB_, bgRGB_);
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
   top_        IN VARCHAR2 := 'thin',
   bottom_     IN VARCHAR2 := 'thin',
   left_       IN VARCHAR2 := 'thin',
   right_      IN VARCHAR2 := 'thin',
   rgb_top_    IN VARCHAR2 := '',
   rgb_bottom_ IN VARCHAR2 := '',
   rgb_left_   IN VARCHAR2 := '',
   rgb_right_  IN VARCHAR2 := '' ) RETURN PLS_INTEGER
IS
   ix_ PLS_INTEGER;
BEGIN
   IF wb_.borders.count > 0 THEN
      FOR b_ IN 0 .. wb_.borders.count - 1 LOOP
         IF (   nvl(wb_.borders(b_).top.style,    'x') = nvl(top_, 'x')
            AND nvl(wb_.borders(b_).top.rgb,      'x') = nvl(rgb_top_, 'x')
            AND nvl(wb_.borders(b_).bottom.style, 'x') = nvl(bottom_, 'x')
            AND nvl(wb_.borders(b_).bottom.rgb,   'x') = nvl(rgb_bottom_, 'x')
            AND nvl(wb_.borders(b_).left.style,   'x') = nvl(left_, 'x')
            AND nvl(wb_.borders(b_).left.rgb,     'x') = nvl(rgb_left_, 'x')
            AND nvl(wb_.borders(b_).right.style,  'x') = nvl(right_, 'x')
            AND nvl(wb_.borders(b_).right.rgb,    'x') = nvl(rgb_right_, 'x')
         ) THEN
            RETURN b_;
         END IF;
      END LOOP;
   END IF;
   ix_ := wb_.borders.count;
   wb_.borders(ix_).top.style    := top_;
   wb_.borders(ix_).top.rgb      := rgb_top_;
   wb_.borders(ix_).bottom.style := bottom_;
   wb_.borders(ix_).bottom.rgb   := rgb_bottom_;
   wb_.borders(ix_).left.style   := left_;
   wb_.borders(ix_).left.rgb     := rgb_left_;
   wb_.borders(ix_).right.style  := right_;
   wb_.borders(ix_).right.rgb    := rgb_right_;
   RETURN ix_;
END Get_Border;

PROCEDURE Get_Border (
   top_        IN VARCHAR2 := 'thin',
   bottom_     IN VARCHAR2 := 'thin',
   left_       IN VARCHAR2 := 'thin',
   right_      IN VARCHAR2 := 'thin',
   rgb_top_    IN VARCHAR2 := '',
   rgb_bottom_ IN VARCHAR2 := '',
   rgb_left_   IN VARCHAR2 := '',
   rgb_right_  IN VARCHAR2 := '' )
IS
   throw_ NUMBER;
BEGIN
   throw_ := Get_Border (top_, bottom_, left_, right_, rgb_top_, rgb_bottom_, rgb_left_, rgb_right_);
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
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   top_        IN VARCHAR2    := '',
   bottom_     IN VARCHAR2    := '',
   left_       IN VARCHAR2    := '',
   right_      IN VARCHAR2    := '',
   rgb_top_    IN VARCHAR2    := '',
   rgb_bottom_ IN VARCHAR2    := '',
   rgb_left_   IN VARCHAR2    := '',
   rgb_right_  IN VARCHAR2    := '',
   sheet_      IN PLS_INTEGER := null )
IS
   sh_          PLS_INTEGER     := nvl (sheet_, wb_.sheets.count);
   Xf_          tp_Xf_fmt       := Get_Cell_Xff (sh_, col_, row_);
   cell_border_ tp_cell_borders := wb_.borders (Xf_.borderId);
   cell_dt_     VARCHAR2(30)    := wb_.sheets(sh_).rows(row_)(col_).datatype;
   border_id_   PLS_INTEGER;
BEGIN

   cell_border_.top.style    := nvl (top_,        cell_border_.top.style);
   cell_border_.top.rgb      := nvl (rgb_top_,    cell_border_.top.rgb);
   cell_border_.bottom.style := nvl (bottom_,     cell_border_.bottom.style);
   cell_border_.bottom.rgb   := nvl (rgb_bottom_, cell_border_.bottom.rgb);
   cell_border_.left.style   := nvl (left_,       cell_border_.left.style);
   cell_border_.left.rgb     := nvl (rgb_left_,   cell_border_.left.rgb);
   cell_border_.right.style  := nvl (right_,      cell_border_.right.style);
   cell_border_.right.rgb    := nvl (rgb_right_,  cell_border_.right.rgb);
   border_id_ := Get_Border (
      cell_border_.top.style, cell_border_.bottom.style, cell_border_.left.style, cell_border_.right.style,
      cell_border_.top.rgb, cell_border_.bottom.rgb, cell_border_.left.rgb, cell_border_.right.rgb
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
   col_start_ IN PLS_INTEGER,
   row_start_ IN PLS_INTEGER,
   col_end_   IN PLS_INTEGER,
   row_end_   IN PLS_INTEGER,
   style_     IN VARCHAR2    := 'medium', -- thin|medium|thick|dotted...
   rgb_       IN VARCHAR2    := '',
   sheet_     IN PLS_INTEGER := null )
IS
   sh_     PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
   width_  PLS_INTEGER := col_end_ - col_start_ + 1;
   height_ PLS_INTEGER := row_end_ - row_start_ + 1;
BEGIN

   -- first we should catch any invalid parameter combinations
   IF width_ < 1 OR height_ < 1 THEN
      Raise_App_Error ('Width and height of a border-range must be greater than zero');

   -- for a 1 x 1 span...
   ELSIF width_ = 1 AND height_ = 1 THEN
      Add_Border_To_Cell (
         col_start_, row_start_, style_, style_, style_, style_,
         rgb_, rgb_, rgb_, rgb_, sh_
      );

   -- for a n x 1 span...
   ELSIF height_ = 1 THEN
      Add_Border_To_Cell (
         col_start_, row_start_, style_, style_, style_, '', rgb_, rgb_, rgb_, '', sh_
      );
      FOR col_ IN (col_start_+1) .. (col_end_-1) LOOP
         Add_Border_To_Cell (
            col_, row_start_, style_, style_, '', '', rgb_, rgb_, '', '', sh_
         );
      END LOOP;
      Add_Border_To_Cell (
         col_end_, row_start_, style_, style_, '', style_, rgb_, rgb_, '', rgb_, sh_
      );

   -- for a 1 x n span
   ELSIF width_ = 1 THEN
      Add_Border_To_Cell (
         col_start_, row_start_, style_, '', style_, style_, rgb_, '', rgb_, rgb_, sh_
      );
      FOR row_ IN (row_start_+1) .. (row_end_-1) LOOP
         Add_Border_To_Cell (
            col_start_, row_, '', '', style_, style_, '', '', rgb_, rgb_, sh_
         );
      END LOOP;
      Add_Border_To_Cell (
         col_start_, row_end_, '', style_, style_, style_, '', rgb_, rgb_, rgb_, sh_
      );

   -- for an n x m span
   ELSE

      FOR col_ IN col_start_ .. col_end_ LOOP
         FOR row_ IN row_start_ .. row_end_ LOOP

            IF col_ = col_start_ THEN -- first column
               IF row_ = row_start_ THEN
                  Add_Border_To_Cell (col_, row_, style_,'',style_,'', rgb_,'',rgb_,'', sh_); -- top-left
               ELSIF row_ = row_end_ THEN
                  Add_Border_To_Cell (col_, row_, '',style_,style_,'', '',rgb_,rgb_,'', sh_); -- bottom-left
               ELSE
                  Add_Border_To_Cell (col_, row_, '','',style_,'', '','',rgb_,'', sh_); -- left-only
               END IF;
            ELSIF col_ = col_end_ THEN -- last column
               IF row_ = row_start_ THEN
                  Add_Border_To_Cell (col_, row_, style_,'','',style_, rgb_,'','',rgb_, sh_); -- top-right
               ELSIF row_ = row_end_ THEN
                  Add_Border_To_Cell (col_, row_, '',style_,'',style_, '',rgb_,'',rgb_, sh_); -- bottom-right
               ELSE
                  Add_Border_To_Cell (col_, row_, '','','',style_, '','','',rgb_, sh_); -- right-only
               END IF;
            ELSE -- middle columns
               IF row_ = row_start_ THEN
                  Add_Border_To_Cell (col_, row_, style_,'','','', rgb_,'','','', sh_); -- top-only
               ELSIF row_ = row_end_ THEN
                  Add_Border_To_Cell (col_, row_, '',style_,'','', '',rgb_,'','', sh_); -- bottom-only
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
IS BEGIN
   Add_Border_To_Range (
      range_.tl.c, range_.tl.r, range_.br.c, range_.br.r, style_, sheet_
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
   wt_tf_    VARCHAR2(1) := CASE WHEN Xf_.alignment.wrapText THEN 't' ELSE 'f' END;
   md5_hash_ RAW(128)    := Dbms_Crypto.Hash (
      Utl_i18n.String_To_Raw (
         to_char(nvl(Xf_.numFmtId,0)) || '^' || to_char(nvl(Xf_.fontId,0)) || '^' || to_char(nvl(Xf_.fillId,0)) ||
         '^' || to_char(nvl(Xf_.borderId,0)) || '^' || nvl (Xf_.alignment.vertical,'x') || '^' ||
         nvl (Xf_.alignment.horizontal,'x') || '^' || wt_tf_,
         'AL32UTF8'
      ),
      dbms_crypto.hash_md5
   );
BEGIN
   FOR i_ IN 1 .. wb_.cellXfs.count LOOP
      IF wb_.cellXfs(i_).md5 = md5_hash_ THEN
         XfId_ := i_;
         exit;
      END IF;
   END LOOP;
   IF XfId_ IS null THEN -- we didn't find a matching style, so create a new one
      xfId_    := wb_.cellXfs.count + 1;
      Xfi_.md5 := md5_hash_;
      wb_.cellXfs(xfId_) := Xfi_;
   END IF;
   RETURN xfId_;
END Get_Or_Create_XfId;

FUNCTION Get_Or_Create_XfId (
   numFmtId_  IN PLS_INTEGER,
   fontId_    IN PLS_INTEGER,
   fillId_    IN PLS_INTEGER,
   borderId_  IN PLS_INTEGER,
   alignment_ IN tp_alignment ) RETURN PLS_INTEGER
IS
   Xf_ tp_Xf_fmt;
BEGIN
   Xf_.numFmtId  := numFmtId_;
   Xf_.fontId    := fontId_;
   Xf_.fillId    := fillId_;
   Xf_.borderId  := borderId_;
   Xf_.alignment := alignment_;
   RETURN Get_Or_Create_XfId (Xf_);
END Get_Or_Create_XfId;

FUNCTION Get_XfId (
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   fillId_    IN PLS_INTEGER  := null,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null ) RETURN PLS_INTEGER
IS BEGIN
   RETURN Get_Or_Create_XfId (numFmtId_, fontId_, fillId_, borderId_, alignment_);
END Get_XfId;

FUNCTION Get_XfId (
   numFmtName_ IN VARCHAR2 := '',
   fontName_   IN VARCHAR2 := '',
   fillName_   IN VARCHAR2 := '',
   borderName_ IN VARCHAR2 := '',
   alignName_  IN VARCHAR2 := '' ) RETURN PLS_INTEGER
IS BEGIN
   RETURN Get_Or_Create_XfId (
      CASE WHEN numFmtName_ IS NOT null THEN numFmt_(numFmtName_) END,
      CASE WHEN fontName_   IS NOT null THEN fonts_(fontName_)    END,
      CASE WHEN fillName_   IS NOT null THEN fills_(fillName_)    END,
      CASE WHEN borderName_ IS NOT null THEN bdrs_(borderName_)   END,
      CASE WHEN alignName_  IS NOT null THEN align_(alignName_)   END
   );
END Get_XfId;

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
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null )
IS
   sh_ PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN
   wb_.sheets(sh_).rows(row_)(col_).datatype  := CELL_DT_NUMBER_;
   wb_.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => '', num_val => value_, dt_val => null
   );
   wb_.sheets(sh_).rows(row_)(col_).value     := value_;
   wb_.sheets(sh_).rows(row_)(col_).style     := CASE
      WHEN xfId_ IS NOT null THEN xfId_
      ELSE get_XfId (
         sh_, col_, row_, numFmtId_, fontId_, fillId_, borderId_, alignment_
      )
   END;
END Cell;

PROCEDURE Cell ( -- num version overload
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_num_  IN NUMBER,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null )
IS
   fm_ix_ PLS_INTEGER := wb_.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   Cell (
      col_, row_, value_num_,
      CASE WHEN numFmtName_ IS NOT null THEN numFmt_(numFmtName_) END,
      CASE WHEN fontName_   IS NOT null THEN fonts_(fontName_) END,
      CASE WHEN fillName_   IS NOT null THEN fills_(fillName_) END,
      CASE WHEN borderName_ IS NOT null THEN bdrs_(borderName_) END,
      CASE WHEN alignName_  IS NOT null THEN align_(alignName_) END,
      sheet_,
      CASE WHEN xfName_     IS NOT null THEN xf_(xfName_) END
   );
   IF formula_ IS NOT null THEN
      wb_.formulas(fm_ix_) := formula_;
      wb_.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
   END IF;
END Cell;

PROCEDURE CellN ( -- num version explicit
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_num_  IN NUMBER,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null )
IS BEGIN
   Cell (
      col_ => col_, row_ => row_, value_num_ => value_num_, formula_ => formula_,
      numFmtName_ => numFmtName_, fontName_  => fontName_,  fillName_ => fillName_,
      borderName_ => borderName_, alignName_ => alignName_, sheet_ => sheet_,
      xfName_ => xfName_
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
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null )
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
   wb_.sheets(sh_).rows(row_)(col_).style := CASE
      WHEN xfId_ IS not null THEN xfId_
      ELSE get_XfId (
         sh_, col_, row_, numFmtId_, fontId_, fillId_, borderId_, align_
      )
   END;
END Cell;

PROCEDURE Cell ( -- string version overload
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_str_  IN VARCHAR2,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null )
IS
   fm_ix_ PLS_INTEGER := wb_.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   Cell (
      col_, row_, value_str_,
      CASE WHEN numFmtName_ IS NOT null THEN numFmt_(numFmtName_) END,
      CASE WHEN fontName_   IS NOT null THEN fonts_(fontName_)    END,
      CASE WHEN fillName_   IS NOT null THEN fills_(fillName_)    END,
      CASE WHEN borderName_ IS NOT null THEN bdrs_(borderName_)   END,
      CASE WHEN alignName_  IS NOT null THEN align_(alignName_)   END,
      sh_,
      CASE WHEN xfName_     IS NOT null THEN xf_(xfName_) END
   );
   IF formula_ IS NOT null THEN
      wb_.formulas(fm_ix_) := formula_;
      wb_.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
   END IF;
END Cell;

PROCEDURE CellS ( -- string version explicit
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_str_  IN VARCHAR2,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null )
IS BEGIN
   Cell (
      col_ => col_, row_ => row_, value_str_ => value_str_, formula_ => formula_,
      numFmtName_ => numFmtName_, fontName_  => fontName_,  fillName_ => fillName_,
      borderName_ => borderName_, alignName_ => alignName_, sheet_ => sheet_,
      xfName_ => xfName_
   );
END CellS;

-----
-- Date_To_Xl_Nr()
-- Excel thinks that 1900 was a leap-year, meaning that the date 1900-02-29 is
-- valid in Excel.  The rest of the world (in particular, Oracle) knows better
-- and so there will always be a discrepancy and a decision to make should you
-- need to "span" over the 1900-02-28 - 1900-03-01 gap.
--   > In Excel:        1900-03-01 - 1900-02-28 = 2
--   > Everywhere else: 1900-03-01 - 1900-02-28 = 1
-- Our solution is to force Excel to show Oracle's (correct) date calculation.
-- Date 1900-03-01 and after assume that 2 refers to 1900-01-01, while earlier
-- dates assume that 1900-01-01 = 1.
-- Just be aware of all this if you need to do some date-calculations in Excel
-- itself, and that it will lead to discrepancies if you need to match answers
-- with calculations made in Oracle.
-- Just to be clear, this is a Microsoft bug, and this Oracle package does its
-- best to work around it.  Your app may require a different approach.
FUNCTION Date_To_Xl_Nr (
   date_ IN DATE ) RETURN NUMBER
IS
   xl_date_as_num_ NUMBER := date_ - to_date('19000301','YYYYMMDD');
BEGIN
   xl_date_as_num_ := xl_date_as_num_ + CASE
      WHEN xl_date_as_num_ < 0 THEN 60 ELSE 61
   END;
   RETURN xl_date_as_num_;
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
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null )
IS
   num_fmt_id_ PLS_INTEGER := numFmtId_;
   sh_         PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
   new_xfId_   PLS_INTEGER := xfId_;
BEGIN
   wb_.sheets(sh_).rows(row_)(col_).datatype  := CELL_DT_DATE_;
   wb_.sheets(sh_).rows(row_)(col_).ora_value := tp_cell_value (
      str_val => '', num_val => null, dt_val => value_
   );
   wb_.sheets(sh_).rows(row_)(col_).value := Date_To_Xl_Nr(value_);
   IF xfId_ IS null THEN
      IF num_fmt_id_ IS null
         AND not (    wb_.sheets(sh_).col_fmts.exists(col_)
                  AND wb_.sheets(sh_).col_fmts(col_).numFmtId IS not null )
         AND not (    wb_.sheets(sh_).row_fmts.exists(row_)
                  AND wb_.sheets(sh_).row_fmts(row_).numFmtId IS not null )
      THEN
         num_fmt_id_ := get_numFmt(dft_fmt_date_short_);
      END IF;
      new_xfId_ := get_xfId (sh_, col_, row_, num_fmt_id_, fontId_, fillId_, borderId_, alignment_);
   END IF;
   wb_.sheets(sh_).rows(row_)(col_).style := new_xfId_;
END Cell;

PROCEDURE Cell ( -- date version overload
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_dt_   IN DATE,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null )
IS
   fm_ix_ PLS_INTEGER := wb_.formulas.count;
   sh_    PLS_INTEGER := nvl (sheet_, wb_.sheets.count);
BEGIN
   Cell (
      col_, row_, value_dt_,
      CASE WHEN numFmtName_ IS NOT null THEN numFmt_(numFmtName_) END,
      CASE WHEN fontName_   IS NOT null THEN fonts_(fontName_)    END,
      CASE WHEN fillName_   IS NOT null THEN fills_(fillName_)    END,
      CASE WHEN borderName_ IS NOT null THEN bdrs_(borderName_)   END,
      CASE WHEN alignName_  IS NOT null THEN align_(alignName_)   END,
      sheet_,
      CASE WHEN xfName_ IS NOT null THEN xf_(xfName_) END
   );
   IF formula_ IS NOT null THEN
      wb_.formulas(fm_ix_) := formula_;
      wb_.sheets(sh_).rows(row_)(col_).formula_idx := fm_ix_;
   END IF;
END Cell;

PROCEDURE CellD ( -- date version explicit
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   value_dt_   IN DATE,
   formula_    IN VARCHAR2    := '',
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null )
IS BEGIN
   Cell (
      col_ => col_, row_ => row_, value_dt_ => value_dt_, formula_ => formula_,
      numFmtName_ => numFmtName_, fontName_ => fontName_, fillName_ => fillName_,
      borderName_ => borderName_, alignName_ => alignName_, sheet_ => sheet_,
      xfName_ => xfName_
   );
END CellD;

-- Sometimes it's useful to "register" a cell which has no value or formatting
-- of any sort.  Note that a cell that contains the string '' adds an entry to
-- the shared strings, so we must instigate this as a to_number(null).
-- Also, it is possible that we want to assign a format mapping or font to our
-- cell, on the off-chance that it will be used as a user-input field.
PROCEDURE CellB (
   col_       IN PLS_INTEGER,
   row_       IN PLS_INTEGER,
   fillId_    IN PLS_INTEGER,
   borderId_  IN PLS_INTEGER  := null,
   alignment_ IN tp_alignment := null,
   numFmtId_  IN PLS_INTEGER  := null,
   fontId_    IN PLS_INTEGER  := null,
   sheet_     IN PLS_INTEGER  := null,
   xfId_      IN PLS_INTEGER  := null )
IS BEGIN
   Cell (
      col_, row_, value_ => to_number(null), numFmtId_ => numFmtId_,
      fontId_ => fontId_, fillId_ => fillId_, borderId_ => borderId_,
      alignment_ => alignment_, sheet_ => sheet_, xfId_ => xfId_
   );
END CellB;
PROCEDURE CellB ( 
   col_        IN PLS_INTEGER,
   row_        IN PLS_INTEGER,
   fillName_   IN VARCHAR2    := null,
   borderName_ IN VARCHAR2    := null,
   alignName_  IN VARCHAR2    := null,
   numFmtName_ IN VARCHAR2    := null,
   fontName_   IN VARCHAR2    := null,
   sheet_      IN PLS_INTEGER := null,
   xfName_     IN VARCHAR2    := null )
IS BEGIN
   Cell (
      col_, row_, value_num_ => to_number(null), numFmtName_ => numFmtName_,
      fontName_ => fontName_, fillName_ => fillName_, borderName_ => borderName_,
      alignName_ => alignName_, sheet_ => sheet_, xfName_ => xfName_
   );
END CellB;

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
   wb_.sheets(sh_).rows(row_)(col_).style     := Get_XfId (
      sh_, col_, row_, fontId_ => Get_Font('Calibri', theme_ => 10, underline_ => true)
   );
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
   ix_ PLS_INTEGER := wb_.sheets(sh_).comments.comments_list.count + 1;
BEGIN
   wb_.sheets(sh_).comments.comments_list(ix_).row    := row_;
   wb_.sheets(sh_).comments.comments_list(ix_).column := col_;
   wb_.sheets(sh_).comments.comments_list(ix_).text   := dbms_xmlgen.convert(text_);
   wb_.sheets(sh_).comments.comments_list(ix_).author := dbms_xmlgen.convert(author_);
   wb_.sheets(sh_).comments.comments_list(ix_).width  := width_;
   wb_.sheets(sh_).comments.comments_list(ix_).height := height_;
END Comment;

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
   ix_ PLS_INTEGER;
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
   file_chunk_ RAW(32);
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
      Dbms_Lob.createTemporary (img_rec_.img_blob, true);

      Dbms_Lob.Copy (img_rec_.img_blob, img_blob_, Dbms_Lob.lobMaxSize, 1, 1);
      img_rec_.img_hash := hash_;
      file_chunk_ := Dbms_Lob.Substr (img_blob_, 32, 1);

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

         offset_ := 5 + to_number(Utl_Raw.Substr(file_chunk_,5,2), 'xxxx');
         LOOP
            file_chunk_ := Dbms_Lob.Substr (img_blob_, 4, offset_);
            hex_        := substr( rawtohex(file_chunk_),1,4);
            EXIT WHEN hex_ IN ('FFDA', 'FFD9') -- SOS Start of Scan; EOI End Of Image
                   OR substr (hex_, 1, 2) != 'FF';
            IF hex_ IN ('FFD0', 'FFD1', 'FFD2', 'FFD3', 'FFD4', 'FFD5', 'FFD6', 'FFD7', /*RSTn*/ 'FF01' /*TEM*/) THEN
               offset_ := offset_ + 2;
            ELSE
               IF hex_ = 'FFC0' /* SOF0 (Start Of Frame 0) marker*/ THEN
                  hex_ := rawtohex (Dbms_Lob.Substr (img_blob_, 4, offset_+5));
                  img_rec_.width  := to_number (substr(hex_,5), 'xxxx');
                  img_rec_.height := to_number (substr(hex_,1,4), 'xxxx');
                  exit;
               END IF;
               offset_ := offset_ + 2 + to_number (utl_raw.substr(file_chunk_,3,2), 'xxxx');
            END IF;
         END LOOP;
         img_rec_.extension := 'jpeg';

      ELSIF utl_raw.substr (file_chunk_,1,2) = '424D' /* BM */ THEN -- bmp
         Dbms_Output.Put_Line ('file is BMP');
         img_rec_.width     := to_number (Utl_Raw.Reverse(Utl_Raw.Substr(file_chunk_,19,4)), 'XXXXXXXX');
         img_rec_.height    := to_number (Utl_Raw.Reverse(Utl_Raw.Substr(file_chunk_,23,4)), 'XXXXXXXX');
         img_rec_.extension := 'bmp';

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
   wb_.sheets(sh_).drawings.drawings_list(wb_.sheets(sh_).drawings.drawings_list.count+1) := drawing_;

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
   Dbms_Lob.loadFromFile (img_blob_, bfile_, Dbms_Lob.getLength(bfile_));
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
IS
   defined_name_ VARCHAR2(100) := Name_Checker (name_);
BEGIN
   wb_.defined_names(defined_name_) := tp_cell_range (
      range_type   => RANGE_DEFINED_NAME_,
      defined_name => defined_name_,
      sheet_id     => sheet_,
      tl           => tp_cell_loc (c => tl_col_, r => tl_row_, fixc => fix_tlc_, fixr => fix_tlr_),
      br           => tp_cell_loc (c => br_col_, r => br_row_, fixc => fix_brc_, fixr => fix_brr_),
      local_sheet  => localsheet_
   );
END Defined_Name;

PROCEDURE Defined_Name (
   range_ IN tp_cell_range )
IS
   rg_ tp_cell_range := range_;
BEGIN
   IF range_.defined_name IS null THEN
      Raise_App_Error ('Defined name cannot be empty!');
   END IF;
   rg_.range_type := RANGE_DEFINED_NAME_;
   wb_.defined_names(range_.defined_name) := rg_;
END Defined_Name;

FUNCTION Range_From_Defined_Name (
   defined_name_ IN VARCHAR2 ) RETURN tp_cell_range
IS BEGIN
   RETURN wb_.defined_names(defined_name_);
END Range_From_Defined_Name;


FUNCTION Create_Pivot_Cache (
   range_    IN tp_cell_range,
   agg_cols_ IN tp_col_cache_method ) RETURN PLS_INTEGER
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
   FOR ix_ IN 1 .. pivot_axes_.col_agg_fns.count LOOP
      IF pivot_axes_.col_agg_fns(ix_).colid = col_ix_ THEN
         --rtn_ := pivot_axes_.col_agg_fns(ix_).agg_fn;
         rtn_ := 'sum'; -- "count" and "sum" are treated the same way in the cache
      END IF;
   END LOOP;
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

FUNCTION Get_Pivot_Table_Data_Source (
   pivot_id_ IN PLS_INTEGER ) RETURN tp_cell_range
IS
BEGIN
   RETURN wb_.pivot_caches(wb_.pivot_tables(pivot_id_).cache_id).ds_range;
END Get_Pivot_Table_Data_Source;

FUNCTION Add_Pivot_Cache (
   src_data_range_ IN OUT NOCOPY tp_cell_range,
   pivot_axes_     IN tp_pivot_axes ) RETURN PLS_INTEGER
IS
   cols_to_cache_ tp_col_cache_method;
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

   -- We only need to create a new cache if the caller hasn't given us a cache
   -- ID of an existing cache.  We shall assume that the caller is intelligent
   -- enough to know that the new pivot-table can only be created from a cache
   -- built on the same data-source.  No validation or stupidity checks are to
   -- be done here!!
   IF cache_id_ IS null THEN
      cache_id_ := Add_Pivot_Cache (src_data_range_, pivot_axes_);
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


PROCEDURE Set_Table (
   col_start_ PLS_INTEGER,
   col_end_   PLS_INTEGER,
   row_start_ PLS_INTEGER,
   row_end_   PLS_INTEGER,
   style_     VARCHAR2,
   tbl_name_  VARCHAR2    := null,
   sheet_     PLS_INTEGER := null )
IS
   table_id_   PLS_INTEGER   := wb_.tables_list.count + 1;
   sh_         PLS_INTEGER   := nvl(sheet_, wb_.sheets.count);
   table_name_ VARCHAR2(100) := Name_Checker (tbl_name_, RANGE_TABLE_);
BEGIN
   IF col_start_ IS null OR col_end_ IS null OR row_start_ IS null OR row_end_ IS null OR sh_ IS null THEN
      Raise_App_Error ('A table''s range must be defined correctly, with full sheet and cell range values.');
   END IF;
   wb_.tables_list(table_id_) := table_name_;
   wb_.defined_names(table_name_) := tp_cell_range (
      range_type   => RANGE_TABLE_,
      defined_name => table_name_,
      sheet_id     => sh_,
      tl           => tp_cell_loc (col_start_, row_start_, false, false),
      br           => tp_cell_loc (col_end_, row_end_, false, false),
      style        => style_
   );
   Add_Col_Headings_To_Range (wb_.defined_names(table_name_), allow_dup_ => false);
   wb_.sheets(sh_).tables_list(table_id_) := table_name_;
END Set_Table;

PROCEDURE Set_Table (
   tbl_range_ tp_cell_range,
   style_     VARCHAR2,
   tbl_name_  VARCHAR2 := null )
IS BEGIN
   Set_Table (
      col_start_ => tbl_range_.tl.c,
      col_end_   => tbl_range_.br.c,
      row_start_ => tbl_range_.tl.r,
      row_end_   => tbl_range_.br.r,
      style_     => style_,
      tbl_name_  => tbl_name_,
      sheet_     => tbl_range_.sheet_id
   );
END Set_Table;


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
   attrs_     nyce_xml.xml_attrs_arr;
   img_exts_  tp_strings;
   ext_       VARCHAR2(5);
   pt_        PLS_INTEGER;
BEGIN

   -- [Content_Types].xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types', attrs_);
   nd_types_ := Nyce_Xml.Make_Root_Node (doc_, 'Types', attrs_);

   IF wb_.images.count > 0 THEN
      FOR img_ IN wb_.images.first .. wb_.images.last LOOP
         ext_ := wb_.images(img_).extension;
         IF ext_ IS NOT null AND not img_exts_.exists(ext_) THEN
            nyce_xml.natr ('ContentType', 'image/' || ext_, attrs_);
            nyce_xml.attr ('Extension', ext_, attrs_);
            Nyce_Xml.Xml_Node (doc_, nd_types_, 'Default', attrs_);
            img_exts_(ext_) := 1;
         END IF;
      END LOOP;
   END IF;

   nyce_xml.natr ('Extension',   'rels', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-package.relationships+xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Default', attrs_);

   nyce_xml.natr ('Extension',   'xml', attrs_);
   nyce_xml.attr ('ContentType', 'application/xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Default', attrs_);

   nyce_xml.natr ('Extension',   'vml', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.vmlDrawing', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Default', attrs_);

   nyce_xml.natr ('PartName', '/xl/workbook.xml', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);

   FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
      nyce_xml.natr ('PartName',    rep('/xl/pivotCache/pivotCacheDefinition:P1.xml', pc_), attrs_);
      nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
      nyce_xml.natr ('PartName', rep('/xl/pivotCache/pivotCacheRecords:P1.xml', pc_), attrs_);
      nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
   END LOOP;
   FOR pt_ IN 1 .. wb_.pivot_tables.count LOOP
      nyce_xml.natr ('PartName', rep('/xl/pivotTables/pivotTable:P1.xml', pt_), attrs_);
      nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
   END LOOP;

   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      nyce_xml.natr ('PartName', rep('/xl/worksheets/sheet:P1.xml', to_char(s_)), attrs_);
      nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
      s_ := wb_.sheets.next(s_);
   END LOOP;

   nyce_xml.natr ('PartName', '/xl/theme/theme1.xml', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.theme+xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
   nyce_xml.natr ('PartName', '/xl/styles.xml', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
   nyce_xml.natr ('PartName', '/xl/sharedStrings.xml', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);

   nyce_xml.natr ('PartName', '/docProps/core.xml', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-package.core-properties+xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
   nyce_xml.natr ('PartName', '/docProps/app.xml', attrs_);
   nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.extended-properties+xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);

   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      IF wb_.sheets(s_).comments.comments_list.count > 0 THEN
         nyce_xml.natr ('PartName', rep('/xl/comments:P1.xml', s_), attrs_);
         nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
      END IF;
      IF wb_.sheets(s_).drawings.drawings_list.count > 0 THEN
         nyce_xml.natr ('PartName', rep('/xl/drawings/drawing:P1.xml', s_), attrs_);
         nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
      END IF;
      s_ := wb_.sheets.next(s_);
   END LOOP;

   FOR t_ IN 1 .. wb_.tables_list.count LOOP
      nyce_xml.natr ('PartName', rep('/xl/tables/table:P1.xml', to_char(t_)), attrs_);
      nyce_xml.attr ('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_types_, 'Override', attrs_);
   END LOOP;

   Add1Xml (excel_, '[Content_Types].xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Content_Types;

PROCEDURE Finish_Rels (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_rels_  dbms_XmlDom.DomNode;
   attrs_    nyce_xml.xml_attrs_arr;
BEGIN

   -- _rels/.rels
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rels_ := Nyce_Xml.Make_Root_Node (doc_, 'Relationships', attrs_);

   nyce_xml.natr ('Id', 'rId1', attrs_);
   nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', attrs_);
   nyce_xml.attr ('Target', 'xl/workbook.xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
   nyce_xml.natr ('Id', 'rId2', attrs_);
   nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', attrs_);
   nyce_xml.attr ('Target', 'docProps/core.xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
   nyce_xml.natr ('Id', 'rId3', attrs_);
   nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', attrs_);
   nyce_xml.attr ('Target', 'docProps/app.xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);

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
   attrs_    nyce_xml.xml_attrs_arr;
BEGIN

   -- docProps/core.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', attrs_);
   nyce_xml.attr ('xmlns:dc', 'http://purl.org/dc/elements/1.1/', attrs_);
   nyce_xml.attr ('xmlns:dcterms', 'http://purl.org/dc/terms/', attrs_);
   nyce_xml.attr ('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', attrs_);
   nyce_xml.attr ('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', attrs_);
   nd_cprop_ := Nyce_Xml.Make_Root_Node (doc_, 'coreProperties', 'cp', attrs_);

   Nyce_Xml.Xml_Text_Node (doc_, nd_cprop_, 'creator',        sys_context('userenv','os_user'), 'dc');
   Nyce_Xml.Xml_Text_Node (doc_, nd_cprop_, 'description',    rep('Build by version: :P1', VERSION_), 'dc');
   Nyce_Xml.Xml_Text_Node (doc_, nd_cprop_, 'lastModifiedBy', sys_context('userenv','os_user'), 'cp');

   nyce_xml.natr ('xsi:type', 'dcterms:W3CDTF', attrs_);
   Nyce_Xml.Xml_Text_Node (doc_, nd_cprop_, 'created',  to_char(current_timestamp,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM'), 'dcterms', attrs_);
   Nyce_Xml.Xml_Text_Node (doc_, nd_cprop_, 'modified', to_char(current_timestamp,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM'), 'dcterms', attrs_);

   Add1Xml (excel_, 'docProps/core.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);


   -- docProps/app.xml
   doc_ := Dbms_XmlDom.newDomDocument;
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties', attrs_);
   nyce_xml.attr ('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes', attrs_);
   nd_prop_ := Nyce_Xml.Make_Root_Node (doc_, 'Properties', attrs_);

   Nyce_Xml.Xml_Text_Node (doc_, nd_prop_, 'Application', 'Microsoft Excel');
   Nyce_Xml.Xml_Text_Node (doc_, nd_prop_, 'DocSecurity', '0');
   Nyce_Xml.Xml_Text_Node (doc_, nd_prop_, 'ScaleCrop', 'false');
   nd_hd_  := Nyce_Xml.Xml_Node (doc_, nd_prop_, 'HeadingPairs');

   nyce_xml.natr ('size',     '2', attrs_);
   nyce_xml.attr ('baseType', 'variant', attrs_);
   nd_vec_ := Nyce_Xml.Xml_Node (doc_, nd_hd_, 'vector', 'vt', attrs_);
   nd_var_ := Nyce_Xml.Xml_Node (doc_, nd_vec_, 'variant', 'vt');
   Nyce_Xml.Xml_Text_Node (doc_, nd_var_, 'lpstr', 'Worksheets', 'vt');
   nd_var_ := Nyce_Xml.Xml_Node (doc_, nd_vec_, 'variant', 'vt');
   Nyce_Xml.Xml_Text_Node (doc_, nd_var_, 'i4', to_char(wb_.sheets.count), 'vt');

   nd_top_ := Nyce_Xml.Xml_Node (doc_, nd_prop_, 'TitlesOfParts');
   nyce_xml.natr ('size', wb_.sheets.count, attrs_);
   nyce_xml.attr ('baseType', 'lpstr', attrs_);
   nd_vec_ := Nyce_Xml.Xml_Node (doc_, nd_top_, 'vector', 'vt', attrs_);
   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      Nyce_Xml.Xml_Text_Node (doc_, nd_vec_, 'lpstr', wb_.sheets(s_).sheet_name, 'vt');
      s_ := wb_.sheets.next(s_);
   END LOOP;
   Nyce_Xml.Xml_Text_Node (doc_, nd_prop_, 'LinksUpToDate', 'false');
   Nyce_Xml.Xml_Text_Node (doc_, nd_prop_, 'SharedDoc', 'false');
   Nyce_Xml.Xml_Text_Node (doc_, nd_prop_, 'HyperlinksChanged', 'false');
   Nyce_Xml.Xml_Text_Node (doc_, nd_prop_, 'AppVersion', '14.0300');

   Add1Xml (excel_, 'docProps/app.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_docProps;

PROCEDURE Finish_Shared_Strings (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_    dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_sst_ dbms_XmlDom.DomNode;
   attrs_  nyce_xml.xml_attrs_arr;
BEGIN

   -- xl/sharedStrings.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   nyce_xml.attr ('count', to_char(wb_.str_cnt), attrs_);
   nyce_xml.attr ('uniqueCount', wb_.strings.count, attrs_);
   nd_sst_ := Nyce_Xml.Make_Root_Node (doc_, 'sst', attrs_);

   nyce_xml.natr ('xml:space', 'preserve', attrs_);
   FOR str_ix_ IN 0 .. wb_.str_ind.count - 1 LOOP
      Nyce_Xml.Xml_Text_Node (doc_, nd_sst_, 'si/t', wb_.str_ind(str_ix_), attrs_);
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
   nd_bdrs_     dbms_XmlDom.DomNode;
   nd_bdr_      dbms_XmlDom.DomNode;
   nd_pf_       dbms_XmlDom.DomNode;
   nd_sxfs_     dbms_XmlDom.DomNode;
   nd_xfs_      dbms_XmlDom.DomNode;
   nd_xf_       dbms_XmlDom.DomNode;
   attrs_       nyce_xml.xml_attrs_arr;

   PROCEDURE Border_Side_Tag (
      nd_parent_ IN dbms_XmlDom.DomNode,
      border_    IN tp_border,
      side_      IN VARCHAR2 )
   IS
      atr_     nyce_xml.xml_attrs_arr;
      nd_side_ dbms_XmlDom.DomNode;
   BEGIN
      IF border_.style IS NOT null THEN
         nyce_xml.attr ('style', border_.style, atr_);
      END IF;
      nd_side_ := Nyce_Xml.Xml_Node (doc_, nd_parent_, side_, atr_);
      IF border_.rgb IS NOT null THEN
         nyce_xml.natr ('rgb', border_.rgb, atr_);
         Nyce_Xml.Xml_Node (doc_, nd_side_, 'color', atr_);
      END IF;
   END Border_Side_Tag;

BEGIN

   -- xl/styles.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.attr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   nyce_xml.attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
   nyce_xml.attr ('mc:Ignorable', 'x14ac', attrs_);
   nyce_xml.attr ('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', attrs_);
   nd_stl_ := Nyce_Xml.Make_Root_Node (doc_, 'styleSheet', attrs_);

   IF wb_.numFmts.count > 0 THEN
      nyce_xml.natr ('count', to_char(wb_.numFmts.count), attrs_);
      nd_numf_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'numFmts', attrs_);
      format_mask_ := wb_.numFmts.first;
      WHILE format_mask_ IS NOT null LOOP
         nyce_xml.natr ('numFmtId', wb_.numFmts(format_mask_), attrs_);
         nyce_xml.attr ('formatCode', format_mask_, attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_numf_, 'numFmt', attrs_);
         format_mask_ := wb_.numFmts.next(format_mask_);
      END LOOP;
   END IF;

   nyce_xml.natr ('count', wb_.fonts.count, attrs_);
   nyce_xml.attr ('x14ac:knownFonts', '1', attrs_);
   nd_fnts_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'fonts', attrs_);
   FOR f_ IN 0 .. wb_.fonts.count-1 LOOP
      nd_fnt_ := Nyce_Xml.Xml_Node (doc_, nd_fnts_, 'font');
      IF wb_.fonts(f_).bold     THEN Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'b'); END IF;
      IF wb_.fonts(f_).italic   THEN Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'i'); END IF;
      IF wb_.fonts(f_).underline THEN Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'u'); END IF;

      nyce_xml.natr ('val', to_char(wb_.fonts(f_).fontsize, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'sz', attrs_);
      nyce_xml.catr (attrs_);
      IF wb_.fonts(f_).rgb IS NOT null THEN
         nyce_xml.attr ('rgb', wb_.fonts(f_).rgb, attrs_);
      ELSE
         nyce_xml.attr ('theme', wb_.fonts(f_).theme, attrs_);
      END IF;
      Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'color', attrs_);

      nyce_xml.natr ('val', wb_.fonts(f_).name, attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'name', attrs_);
      nyce_xml.natr ('val', wb_.fonts(f_).family, attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'family', attrs_);
      nyce_xml.natr ('val', 'none', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_fnt_, 'scheme', attrs_);
   END LOOP;

   nyce_xml.natr ('count', wb_.fills.count, attrs_);
   nd_fills_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'fills', attrs_);
   FOR f_ IN 0 .. wb_.fills.count-1 LOOP
      nyce_xml.natr ('patternType', wb_.fills(f_).patternType, attrs_);
      nd_pf_ := Nyce_Xml.Xml_Node (doc_, nd_fills_, 'fill/patternFill', attrs_);
      nyce_xml.catr (attrs_);
      IF wb_.fills(f_).fgRGB IS NOT null THEN
         nyce_xml.attr ('rgb', wb_.fills(f_).fgRGB, attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_pf_, 'fgColor', attrs_);
      END IF;
      IF wb_.fills(f_).bgRGB IS NOT null THEN
         nyce_xml.attr ('rgb', wb_.fills(f_).bgRGB, attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_pf_, 'bgColor', attrs_);
      END IF;
   END LOOP;

   nyce_xml.natr ('count', wb_.borders.count, attrs_);
   nd_bdrs_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'borders', attrs_);
   FOR b_ IN 0 .. wb_.borders.count-1 LOOP
      nd_bdr_ := Nyce_Xml.Xml_Node (doc_, nd_bdrs_, 'border');
      Border_Side_Tag (nd_bdr_, wb_.borders(b_).left,   'left');
      Border_Side_Tag (nd_bdr_, wb_.borders(b_).right,  'right');
      Border_Side_Tag (nd_bdr_, wb_.borders(b_).top,    'top');
      Border_Side_Tag (nd_bdr_, wb_.borders(b_).bottom, 'bottom');
   END LOOP;

   nyce_xml.natr ('count', '1', attrs_);
   nd_sxfs_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'cellStyleXfs', attrs_);
   nyce_xml.natr ('numFmtId', '0', attrs_);
   nyce_xml.attr ('fontId', '0', attrs_);
   nyce_xml.attr ('fillId', '0', attrs_);
   nyce_xml.attr ('borderId', '0', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_sxfs_, 'xf', attrs_);

   nyce_xml.natr ('count', wb_.cellXfs.count+1, attrs_);
   nd_xfs_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'cellXfs', attrs_);

   nyce_xml.natr ('numFmtId', '0', attrs_);
   nyce_xml.attr ('fontId', '0', attrs_);
   nyce_xml.attr ('fillId', '0', attrs_);
   nyce_xml.attr ('borderId', '0', attrs_);
   nyce_xml.attr ('xfId', '0', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_xfs_, 'xf', attrs_);
   FOR x_ IN 1 .. wb_.cellXfs.count LOOP
      nyce_xml.catr (attrs_);
      nyce_xml.natr ('numFmtId', wb_.cellXfs(x_).numFmtId, attrs_);
      nyce_xml.attr ('fontId', wb_.cellXfs(x_).fontId, attrs_);
      nyce_xml.attr ('fillId', wb_.cellXfs(x_).fillId, attrs_);
      nyce_xml.attr ('borderId', wb_.cellXfs(x_).borderId, attrs_);
      nd_xf_ := Nyce_Xml.Xml_Node (doc_, nd_xfs_, 'xf', attrs_);
      IF wb_.cellXfs(x_).alignment.horizontal IS NOT null OR wb_.cellXfs(x_).alignment.vertical IS NOT null OR wb_.cellXfs(x_).alignment.wrapText IS NOT null THEN
         nyce_xml.catr (attrs_);
         IF wb_.cellXfs(x_).alignment.horizontal IS NOT null THEN nyce_xml.attr('horizontal', wb_.cellXfs(x_).alignment.horizontal, attrs_); END IF;
         IF wb_.cellXfs(x_).alignment.vertical    IS NOT null THEN nyce_xml.attr('vertical', wb_.cellXfs(x_).alignment.vertical, attrs_); END IF;
         IF wb_.cellXfs(x_).alignment.wrapText THEN nyce_xml.attr('wrapText', 'true', attrs_); END IF;
         Nyce_Xml.Xml_Node (doc_, nd_xf_, 'alignment', attrs_);
      END IF;
   END LOOP;

   nyce_xml.natr ('count', '1', attrs_);
   nd_xfs_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'cellStyles', attrs_);

   nyce_xml.natr ('name', 'Normal', attrs_);
   nyce_xml.attr ('xfId', '0', attrs_);
   nyce_xml.attr ('builtinId', '0', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_xfs_, 'cellStyle', attrs_);

   nyce_xml.natr ('count', '0', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_stl_, 'dxfs', attrs_);
   nyce_xml.natr ('defaultTableStyle', 'TableStyleMedium2', attrs_);
   nyce_xml.attr ('defaultPivotStyle', 'PivotStyleLight16', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_stl_, 'tableStyles', attrs_);

   nd_xfs_ := Nyce_Xml.Xml_Node (doc_, nd_stl_, 'extLst');
   nyce_xml.natr ('uri', '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}', attrs_);
   nyce_xml.attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
   nd_xf_ := Nyce_Xml.Xml_Node (doc_, nd_xfs_, 'ext', attrs_);
   nyce_xml.natr ('defaultSlicerStyle', 'SlicerStyleLight1', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_xf_, 'slicerStyles', 'x14', attrs_);

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
   attrs_    nyce_xml.xml_attrs_arr;
   s_        PLS_INTEGER;
   dn_       VARCHAR2(100);
   rel_      PLS_INTEGER := 4; -- see hard-coded rels in Finish_Workbook_Rels()
BEGIN

   -- xl/workbook.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   nyce_xml.attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
   nd_wb_ := Nyce_Xml.Make_Root_Node (doc_, 'workbook', attrs_);

   nyce_xml.natr ('appName', 'xl', attrs_);
   nyce_xml.attr ('lastEdited', '5', attrs_);
   nyce_xml.attr ('lowestEdited', '5', attrs_);
   nyce_xml.attr ('rupBuild', '9302', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_wb_, 'fileVersion', attrs_);
   nyce_xml.catr (attrs_);
   IF wb_.pivot_tables.count > 0 THEN
      nyce_xml.attr ('hidePivotFieldList', '1', attrs_);
   END IF;
   nyce_xml.attr ('defaultThemeVersion', '166925', attrs_);
   nyce_xml.attr ('date1904', 'false', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_wb_, 'workbookPr', attrs_);

   nd_bks_ := Nyce_Xml.Xml_Node (doc_, nd_wb_, 'bookViews');
   nyce_xml.natr ('xWindow',  '120', attrs_);
   nyce_xml.attr ('yWindow', '45', attrs_);
   nyce_xml.attr ('windowWidth', '19155', attrs_);
   nyce_xml.attr ('windowHeight', '4935', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_bks_, 'workbookView', attrs_);

   nd_shs_ := Nyce_Xml.Xml_Node (doc_, nd_wb_, 'sheets');
   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      nyce_xml.natr ('name', wb_.sheets(s_).sheet_name, attrs_);
      nyce_xml.attr ('sheetId', to_char(s_), attrs_);
      nyce_xml.attr ('r:id', rep ('rId:P1', to_char(rel_)), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_shs_, 'sheet', attrs_);
      wb_.sheets(s_).wb_rel := rel_;
      rel_ := rel_ + 1;
      s_   := wb_.sheets.next(s_);
   END LOOP;

   IF wb_.defined_names.count - wb_.tables_list.count > 0 THEN
      nd_dnm_ := Nyce_Xml.Xml_Node (doc_, nd_wb_, 'definedNames');
      dn_ := wb_.defined_names.first;
      WHILE dn_ IS NOT null LOOP
         IF wb_.defined_names(dn_).range_type = RANGE_DEFINED_NAME_ THEN
            nyce_xml.natr ('name', dn_, attrs_);
            IF wb_.defined_names(dn_).local_sheet THEN
               IF wb_.defined_names(dn_).sheet_id IS null THEN
                  Raise_App_Error ('Sheet Id must be defined for local-sheet function to be viable!');
               END IF;
               nyce_xml.attr ('localSheetId', to_char(wb_.defined_names(dn_).sheet_id), attrs_);
            END IF;
            Nyce_Xml.Xml_Text_Node (doc_, nd_dnm_, 'definedName', Alfan_Sheet_Range(wb_.defined_names(dn_)), attrs_);
         END IF;
         dn_ := wb_.defined_names.next(dn_);
      END LOOP;
   END IF;

   nyce_xml.natr ('calcId', '144525', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_wb_, 'calcPr', attrs_);

   IF wb_.pivot_caches.count > 0 THEN
      nd_pvs_ :=  Nyce_Xml.Xml_Node (doc_, nd_wb_, 'pivotCaches');
      FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
         nyce_xml.natr ('cacheId', to_char(wb_.pivot_caches(pc_).cache_id), attrs_);
         nyce_xml.attr ('r:id', 'rId' || to_char(rel_), attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_pvs_, 'pivotCache', attrs_);
         wb_.pivot_caches(pc_).wb_rel := rel_;
         rel_                         := rel_ + 1;
      END LOOP;

      nd_extl_ := Nyce_Xml.Xml_Node (doc_, nd_wb_, 'extLst');

      nyce_xml.natr ('uri', Get_Guid, attrs_);
      nyce_xml.attr ('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main', attrs_);
      nd_ext_ := Nyce_Xml.Xml_Node (doc_, nd_extl_, 'ext', attrs_);

      nyce_xml.natr ('chartTrackingRefBase', '1', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ext_, 'workbookPr', 'x15', attrs_);

      nyce_xml.natr ('uri', Get_Guid, attrs_);
      nyce_xml.attr ('xmlns:xcalcf', 'http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures', attrs_);
      nd_ext_ := Nyce_Xml.Xml_Node (doc_, nd_extl_, 'ext', attrs_);

      nd_cf_ := Nyce_Xml.Xml_Node (doc_, nd_ext_, 'calcFeatures', 'xcalcf');

      nyce_xml.natr ('name', 'microsoft.com:RD', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      nyce_xml.natr ('name', 'microsoft.com:Single', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      nyce_xml.natr ('name', 'microsoft.com:FV', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      nyce_xml.natr ('name', 'microsoft.com:CNMTM', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);
      nyce_xml.natr ('name', 'microsoft.com:LET_WF', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_cf_, 'feature', 'xcalcf', attrs_);

   END IF;

   Add1Xml (excel_, 'xl/workbook.xml', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Workbook;

PROCEDURE Finish_Workbook_Rels (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_    dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_rls_ dbms_XmlDom.DomNode;
   attrs_  nyce_xml.xml_attrs_arr;
   s_      PLS_INTEGER;
BEGIN

   -- xl/_rels/workbook.xml.rels
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rls_ := Nyce_Xml.Make_Root_Node (doc_, 'Relationships', attrs_);

   nyce_xml.natr ('Id', 'rId1', attrs_);
   nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', attrs_);
   nyce_xml.attr ('Target', 'sharedStrings.xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);

   nyce_xml.natr ('Id', 'rId2', attrs_);
   nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', attrs_);
   nyce_xml.attr ('Target', 'styles.xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);

   nyce_xml.natr ('Id', 'rId3', attrs_);
   nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', attrs_);
   nyce_xml.attr ('Target', 'theme/theme1.xml', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);

   FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
      nyce_xml.natr ('Id', 'rId' || to_char (wb_.pivot_caches(pc_).wb_rel), attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition', attrs_);
      nyce_xml.attr ('Target', rep ('pivotCache/pivotCacheDefinition:P1.xml', pc_), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);
   END LOOP;

   s_ := wb_.sheets.first;
   WHILE s_ IS NOT null LOOP
      nyce_xml.natr ('Id', 'rId' || to_char(wb_.sheets(s_).wb_rel), attrs_);
      nyce_xml.attr ('Type',  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', attrs_);
      nyce_xml.attr ('Target', rep ('worksheets/sheet:P1.xml', to_char(s_)), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rls_, 'Relationship', attrs_);
      s_ := wb_.sheets.next(s_);
   END LOOP;

   Add1Xml (excel_, 'xl/_rels/workbook.xml.rels', Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

END Finish_Workbook_Rels;

PROCEDURE Finish_Media (
   excel_ IN OUT NOCOPY BLOB )
IS BEGIN
   FOR img_ IN 1 .. wb_.images.count LOOP
      Add1File (
         zipped_blob_ => excel_,
         filename_    => rep ('xl/media/image:P1.:P2', img_, wb_.images(img_).extension),
         content_     => wb_.images(img_).img_blob
      );
   END LOOP;
END Finish_Media;

PROCEDURE Build_Pivot_Caches_And_Tables
IS
   rollup_fn_      VARCHAR2(20);
   col_name_       VARCHAR2(32000);
   cache_field_    tp_cache_field;
   min_val_        NUMBER;
   max_val_        NUMBER;
   ord_si_         tp_data_ix_ord;
   uq_si_          tp_unique_data;
BEGIN
   -- Build out the caches with necessary data
   FOR pc_ IN 0 .. wb_.pivot_caches.count-1 LOOP
      FOR c_ IN 1 .. Range_Width (wb_.pivot_caches(pc_).ds_range) LOOP

         col_name_  := Range_Col_Head_Name (wb_.pivot_caches(pc_).ds_range, c_);
         rollup_fn_ := wb_.pivot_caches(pc_).flds_to_cache(c_);
         ord_si_    := tp_data_ix_ord();
         uq_si_     := tp_unique_data();

         IF rollup_fn_ IN ('row','col','filter') THEN -- filter needs to be checked!!
            Build_Si_From_Range (uq_si_, ord_si_, wb_.pivot_caches(pc_).ds_range, c_);
            cache_field_ := tp_cache_field (
               field_name   => col_name_,
               rollup_fn    => rollup_fn_,
               format_id    => Range_Col_NumFmtId (wb_.pivot_caches(pc_).ds_range, c_),
               shared_items => uq_si_,
               si_order     => ord_si_,
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

   -- Then build out the pivot-tables in the same manner
   FOR pt_ IN 1 .. wb_.pivot_tables.count LOOP
      IF wb_.pivot_tables(pt_).pivot_axes.col_agg_fns.count = 1 THEN
         wb_.pivot_tables(pt_).pivot_axes.col_agg_fns(1).col_tot_name := 'Total';
      END IF;
      FOR ag_ IN 1 .. wb_.pivot_tables(pt_).pivot_axes.col_agg_fns.count LOOP
         wb_.pivot_tables(pt_).pivot_axes.col_agg_fns(ag_).col_agg_name := CASE
            wb_.pivot_tables(pt_).pivot_axes.col_agg_fns(ag_).agg_fn
               WHEN 'count' THEN 'Count of '
               WHEN 'sum'   THEN 'Sum of '
         END || Range_Col_Head_Name (
            range_    => Get_Pivot_Table_Data_Source (pt_),
            col_offs_ => wb_.pivot_tables(pt_).pivot_axes.col_agg_fns(ag_).colid
         );
      END LOOP;
      wb_.pivot_tables(pt_).pivot_axes.col_agg_fns(1).col_tot_name := CASE
         WHEN wb_.pivot_tables(pt_).pivot_axes.col_agg_fns.count = 1 THEN 'Total'
         ELSE wb_.pivot_tables(pt_).pivot_axes.col_agg_fns(1).col_agg_name
      END;
   END LOOP;
END Build_Pivot_Caches_And_Tables;

PROCEDURE Finish_Pivot_Caches (
   excel_ IN OUT NOCOPY BLOB )
IS
   sh_      PLS_INTEGER;
   doc_     dbms_XmlDom.DomDocument;
   attrs_   nyce_xml.xml_attrs_arr;
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

      nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
      nyce_xml.attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
      nyce_xml.attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
      nyce_xml.attr ('mc:Ignorable', 'xr', attrs_);
      nyce_xml.attr ('r:id', 'rId1', attrs_); -- points to pivot-cache-record; there's only ever 1 per pCD
      nyce_xml.attr ('refreshedBy', user, attrs_);
      nyce_xml.attr ('refreshedDate', Date_To_Xl_Nr(sysdate), attrs_);
      nyce_xml.attr ('createdVersion', '7', attrs_); -- Version of Excel in which this pivot was created!
      nyce_xml.attr ('refreshedVersion', '7', attrs_);
      nyce_xml.attr ('minRefreshableVersion', '3', attrs_); -- Minimum version of Excel which is compatible (apparently)
      nyce_xml.attr ('recordCount', to_char(Range_Height(cache_.ds_range)), attrs_);
      nyce_xml.attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
      nyce_xml.attr ('xr:uid', Get_Guid, attrs_); --'{C898DCD4-A18D-452F-B655-4FAEB857F78F}';
      nd_pcd_ := Nyce_Xml.Make_Root_Node (doc_, 'pivotCacheDefinition', attrs_);

      nyce_xml.natr ('type', 'worksheet', attrs_);
      nd_cs_ := Nyce_Xml.Xml_Node (doc_, nd_pcd_, 'cacheSource', attrs_);

      nyce_xml.catr (attrs_);
      IF cache_.ds_range.defined_name IS NOT null THEN
         nyce_xml.attr ('name', cache_.ds_range.defined_name, attrs_);
      ELSE
         nyce_xml.attr ('ref', Alfan_Range (cache_.ds_range), attrs_);
         nyce_xml.attr ('sheet', Sheet_Name (cache_.ds_range), attrs_);
      END IF;
      Nyce_Xml.Xml_Node (doc_, nd_cs_, 'worksheetSource', attrs_);

      nyce_xml.catr (attrs_);
      nyce_xml.attr ('count', to_char(cache_.cf_order.count), attrs_);
      nd_cfs_ := Nyce_Xml.Xml_Node (doc_, nd_pcd_, 'cacheFields', attrs_);

      FOR c_ IN cache_.cf_order.first .. cache_.cf_order.last LOOP

         fld_  := cache_.cf_order(c_);
         cfld_ := cache_.cached_fields(fld_);

         nyce_xml.natr ('name', fld_, attrs_);
         nyce_xml.attr ('numFmtId', cache_.cached_fields(fld_).format_id, attrs_);
         nd_cf_ := Nyce_Xml.Xml_Node (doc_, nd_cfs_, 'cacheField', attrs_);

         nyce_xml.catr (attrs_);
         IF cfld_.rollup_fn IN ('row','column','fileter') THEN
            nyce_xml.attr ('count', to_char(cache_.cached_fields(fld_).shared_items.count), attrs_);
         ELSIF cfld_.rollup_fn = 'sum' THEN
            nyce_xml.attr ('containsSemiMixedTypes', '0', attrs_);
            nyce_xml.attr ('containsString', '0', attrs_);
            nyce_xml.attr ('containsNumber', '1', attrs_);
            nyce_xml.attr ('minValue', to_char(cache_.cached_fields(fld_).min_value), attrs_);
            nyce_xml.attr ('maxValue', to_char(cache_.cached_fields(fld_).max_value), attrs_);
         END IF;
         nd_si_ := Nyce_Xml.Xml_Node (doc_, nd_cf_, 'sharedItems', attrs_);

         IF cfld_.si_order.count > 0 THEN
            FOR si_ IN cfld_.si_order.first .. cfld_.si_order.last LOOP
               nyce_xml.natr ('v', cfld_.si_order(si_), attrs_);
               Nyce_Xml.Xml_Node (doc_, nd_si_, 's', attrs_); -- s for a string, which we assume, for now
            END LOOP;
         END IF;
      END LOOP;
      nd_el_ := Nyce_Xml.Xml_Node (doc_, nd_pcd_, 'extLst');

      nyce_xml.natr ('uri', Get_Guid, attrs_);
      nyce_xml.attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
      nd_ex_ := Nyce_Xml.Xml_Node (doc_, nd_el_, 'ext', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ex_, 'pivotCacheDefinition', 'x14');

      Add1Xml (excel_, rep('xl/pivotCache/pivotCacheDefinition:P1.xml',pc_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);


      -- xl/pivotCache/pivotCacheRecords:P1.xml
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
      nyce_xml.attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
      nyce_xml.attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
      nyce_xml.attr ('mc:Ignorable', 'xr', attrs_);
      nyce_xml.attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
      nyce_xml.attr ('count', to_char(cache_.ds_range.br.r - cache_.ds_range.tl.r), attrs_);
      nd_pcd_ := Nyce_Xml.Make_Root_Node (doc_, 'pivotCacheRecords', attrs_);

      FOR r_ IN cache_.ds_range.tl.r+1 .. cache_.ds_range.br.r LOOP
         nd_row_ := Nyce_Xml.Xml_Node (doc_, nd_pcd_, 'r');
         FOR c_ IN cache_.cf_order.first .. cache_.cf_order.last LOOP
            cfld_   := cache_.cached_fields(cache_.cf_order(c_));
            xl_col_ := cache_.ds_range.tl.c + c_ - 1; -- one based, not zero
            nyce_xml.natr ('v', Get_Cell_Cache_Value (xl_col_, r_, sh_, cfld_.shared_items), attrs_);
            tag_    := Get_Cell_Cache_Tag (xl_col_, r_, sh_, cfld_.rollup_fn);
            Nyce_Xml.Xml_Node (doc_, nd_row_, tag_, attrs_);
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

      nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
      nd_rels_ := Nyce_Xml.Make_Root_Node (doc_, 'Relationships', attrs_);

      nyce_xml.natr ('Id', 'rId1', attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords', attrs_);
      nyce_xml.attr ('Target', rep ('pivotCacheRecords:P1.xml', pc_), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);

      Add1Xml (excel_, rep('xl/pivotCache/_rels/pivotCacheDefinition:P1.xml.rels',pc_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);

   END LOOP;

END Finish_Pivot_Caches;

-----
-- Combine_Arrays()
--   Combines interger-indexed arrays of strings
--     -> [1 = 'a', 3 => 'b'] + [2 => 'c'] becomes [1 => 'a', 2 => 'c', 3 => 'b']
--
FUNCTION Combine_Arrays (
   arr1_ IN tp_col_filters,
   arr2_ IN tp_col_filters,
   arr3_ IN tp_col_filters := tp_col_filters(),
   arr4_ IN tp_col_filters := tp_col_filters() ) RETURN tp_col_filters
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
   ix_ := arr3_.first;
   WHILE ix_ IS NOT null LOOP
      IF ret_.exists(ix_) THEN
         Raise_App_Error ('Duplicate indexes :P1 when combining arrays', to_char(ix_));
      END IF;
      ix_ := arr3_.next(ix_);
   END LOOP;
   ix_ := arr4_.first;
   WHILE ix_ IS NOT null LOOP
      IF ret_.exists(ix_) THEN
         Raise_App_Error ('Duplicate indexes :P1 when combining arrays', to_char(ix_));
      END IF;
      ix_ := arr4_.next(ix_);
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
         'Pivot table has expanded into sheet/cell [:P1/:P2] which already contains data.' ||
         '  To avoid any chance of recursive references, this is not allowed.',
         sh_, Alfan_Cell (col_, row_)
      );
   END IF;
END Check_Cell_Not_Exist;


-----
-- Unravel_Json_To_Sheet()
--   Once rolled up, we first need to use the JSON to populate this workbook's
--   sheets.  Dependant on the pivot caches, it may well be that the new pivot
--   tables overwirte existing cells, which raises an exception in our program
--   just as it does in actual Excel.
--
PROCEDURE Unravel_Json_To_Sheet (
   pivot_id_ IN PLS_INTEGER,
   j_piv_    IN json_object_t )
IS
   pt_loc_     tp_cell_loc   := wb_.pivot_tables(pivot_id_).location_tl;
   sh_         PLS_INTEGER   := wb_.pivot_tables(pivot_id_).on_sheet;
   col_        PLS_INTEGER;
   row_        PLS_INTEGER;
   v_row_      PLS_INTEGER;
   hd_arr_     json_array_t  := j_piv_.get_array ('xlPtHead');
   vt_arr_     json_array_t  := j_piv_.get_array ('xlPtvAxes');
   grid_arr_   json_array_t  := j_piv_.get_object('full-grid').get_array('xlGrid');
   lv_arr_     json_array_t;
BEGIN
   -- Paste the pre-calculated header into the Excel sheet
   FOR r_ IN 0 .. hd_arr_.get_size-1 LOOP
      row_    := pt_loc_.r + r_;
      lv_arr_ := treat (hd_arr_.get(r_) as json_array_t);
      FOR c_ IN 0 .. (lv_arr_.get_size-1) LOOP
         col_ := pt_loc_.c + c_;
         Check_Cell_Not_Exist (sh_, col_, row_);
         CellS (col_, row_, lv_arr_.get_string(c_), sheet_ => sh_);
      END LOOP;
   END LOOP;
   -- Then the vertical row-header names
   FOR r_ IN 0 .. vt_arr_.get_size-1 LOOP
      v_row_ := row_ + r_ + 1;
      Check_Cell_Not_Exist (sh_, pt_loc_.c, v_row_);
      CellS (pt_loc_.c, v_row_, treat(vt_arr_.get(r_) as json_object_t).get_string('val'), sheet_ => sh_);
   END LOOP;
   -- And then the pivoted data itself...
   FOR r_ IN 0 .. grid_arr_.get_size-1 LOOP
      v_row_  := row_ + r_ + 1;
      lv_arr_ := treat(grid_arr_.get(r_) as json_object_t).get_array('grid');
      FOR c_ IN 0 .. lv_arr_.get_size-1 LOOP
         col_ := pt_loc_.c + 1 + c_;
         Check_Cell_Not_Exist (sh_, col_, v_row_);
         CellN (col_, v_row_, lv_arr_.get_number(c_), sheet_ => sh_);
      END LOOP;
   END LOOP;
END Unravel_Json_To_Sheet;

-----
-- Unravel_Json_Ptv_Axes_Xml()
-- Unravel_Json_Pth_Axes_Xml()
--   Build the <rowItems> and <colItems> XML tags for PivotTable.xml file-part
--   of the Excel sheet.  Both horizontal and vertical axes have been built as
--   a Json array or object, including the "level" information.
--
PROCEDURE Unravel_Json_Ptv_Axes_Xml (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   xml_nd_   IN            dbms_XmlDom.DomNode,
   axes_arr_ IN            json_array_t,
   cache_    IN            tp_pivot_cache )
IS
   nd_i_     dbms_XmlDom.DomNode;
   attrs_    nyce_xml.xml_attrs_arr;
   ix_obj_   json_object_t;
   lv_       PLS_INTEGER;
   col_name_ VARCHAR2(32000);
   si_val_   VARCHAR2(32000);
   v_        PLS_INTEGER;
BEGIN
   FOR ix_ IN 0 .. axes_arr_.get_size-1 LOOP

      ix_obj_   := treat(axes_arr_.get(ix_) as json_object_t);
      lv_       := ix_obj_.get_number ('lv') - 1; -- 'lv' starts at 1, so lv_=0 is root-level
      col_name_ := ix_obj_.get_string ('colName');
      si_val_   := ix_obj_.get_string ('val');
      v_        := 0;

      IF lv_ >= 0 THEN
         nyce_xml.natr ('r', to_char(lv_), attrs_, lv_>0);
         nd_i_ := Nyce_Xml.Xml_Node (doc_, xml_nd_, 'i', attrs_);
         v_ := cache_.cached_fields(col_name_).shared_items(si_val_);
         nyce_xml.natr ('v', v_, attrs_, v_>0);
         Nyce_Xml.Xml_Node (doc_, nd_i_, 'x', attrs_);

      ELSE -- last record will be -1
         nyce_xml.natr ('t', 'grand', attrs_);
         nd_i_ := Nyce_Xml.Xml_Node (doc_, xml_nd_, 'i', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_i_, 'x');

      END IF;
   END LOOP;
END Unravel_Json_Ptv_Axes_Xml;

PROCEDURE Unravel_Json_Pth_Axes_Xml (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   xml_nd_   IN            dbms_XmlDom.DomNode,
   axes_obj_ IN            json_object_t,
   cache_    IN            tp_pivot_cache )
IS
   col_names_arr_   json_array_t := axes_obj_.get_array('colNames');
   axes_arr_        json_array_t := axes_obj_.get_array('vHead');
   pthd_v_col_      json_array_t;
   grand_total_     BOOLEAN      := false;
   totals_col_name_ VARCHAR2(32000);
   totals_col_val_  VARCHAR2(32000);
   col_name_        VARCHAR2(32000);
   col_val_         VARCHAR2(32000);
   r_val_           PLS_INTEGER;
   attrs_           nyce_xml.xml_attrs_arr;
   nd_i_            dbms_XmlDom.DomNode;
   v_               PLS_INTEGER;

   FUNCTION Is_Total (
      col_val_ IN VARCHAR2 ) RETURN BOOLEAN
   IS BEGIN
      RETURN col_val_ LIKE '% Total' OR col_val_ LIKE 'Count of %' OR col_val_ LIKE 'Sum of %';
   END Is_Total;
   FUNCTION Col_Val_From_Total (
      chk_col_val_ IN VARCHAR2 ) RETURN VARCHAR2
   IS
      new_val_ VARCHAR2(32000);
   BEGIN
      new_val_ := replace (chk_col_val_, ' Total');
      new_val_ := replace (new_val_, 'Count of ');
      new_val_ := replace (new_val_, 'Sum of ');
      RETURN new_val_;
   END Col_Val_From_Total;

BEGIN
   FOR pt_col_ IN 0 .. axes_arr_.get_size-1 LOOP -- loop on each <i>

      nyce_xml.catr (attrs_);
      r_val_           := 0;
      totals_col_name_ := '';
      totals_col_val_  := '';
      pthd_v_col_      := treat (axes_arr_.get(pt_col_) as json_array_t);

      -- This first loop checks for null values and "total" values in the head
      -- grid-array.  If none are found, we process the same loop again below.
      FOR lv_ IN 0 .. pthd_v_col_.get_size-1 LOOP
         col_val_ := pthd_v_col_.get_string(lv_);
         IF col_val_ IS null THEN
            r_val_ := r_val_ + 1;
         ELSIF col_val_ = 'Grand Total' OR col_val_ LIKE 'Total %' THEN
            grand_total_ := true;
            nyce_xml.attr ('t', 'grand', attrs_);
            exit;
         ELSIF Is_Total(col_val_) THEN
            totals_col_name_ := col_names_arr_.get_string(lv_);
            totals_col_val_  := Col_Val_From_Total (col_val_);
            nyce_xml.attr ('t', 'default', attrs_);
            exit;
         END IF;
      END LOOP;

      -- create <i> node
      nyce_xml.attr ('r', to_char(r_val_), attrs_, r_val_>0);
      nd_i_ := Nyce_Xml.Xml_Node (doc_, xml_nd_, 'i', attrs_);

      -- <i><x> nodes...
      IF grand_total_ THEN
         Nyce_Xml.Xml_Node (doc_, nd_i_, 'x');

      ELSIF totals_col_name_ IS NOT null THEN
         v_ := cache_.cached_fields(totals_col_name_).shared_items(totals_col_val_);
         nyce_xml.natr ('v', to_char(v_), attrs_, v_>0);
         Nyce_Xml.Xml_Node (doc_, nd_i_, 'x', attrs_);

      ELSE
         -- There shouldn't be any more totals here as we've already processed
         -- the array up to the totals in the earlier loop
         FOR lv_ IN r_val_ .. pthd_v_col_.get_size-1 LOOP
            col_name_ := col_names_arr_.get_string(lv_);
            col_val_  := pthd_v_col_.get_string(lv_);
            v_ := cache_.cached_fields(col_name_).shared_items(col_val_);
            nyce_xml.natr ('v', to_char(v_), attrs_, v_>0);
            Nyce_Xml.Xml_Node (doc_, nd_i_, 'x', attrs_);
         END LOOP;
      END IF;

   END LOOP;
END Unravel_Json_Pth_Axes_Xml;


-----
-- Initiate_Sum_Object()
--   We need to initiate and update the sums object, which is a bit of a faff!
--   It's because technically, a pivot table can calculate sums, averages, and
--   many more types of aggregates over several columns.
--   Only 'sum' is supported at the moment!!
--
PROCEDURE Initiate_Agg_Obj (
   aggs_rec_ IN            tp_col_agg_fns,
   aggs_obj_ IN OUT NOCOPY json_object_t )
IS
   agg_obj_    json_object_t := json_object_t();
   aggs_list_  json_array_t  := json_array_t();
BEGIN
   aggs_obj_ := json_object_t();
   FOR ix_ IN 1 .. aggs_rec_.count LOOP
      agg_obj_.put ('colId',   aggs_rec_(ix_).colid);
      agg_obj_.put ('colName', aggs_rec_(ix_).col_tot_name);
      agg_obj_.put ('fn',      aggs_rec_(ix_).agg_fn);
      agg_obj_.put ('value',   to_number(null));
      aggs_list_.append(agg_obj_);
   END LOOP;
   aggs_obj_.put ('recordCount', 0);
   aggs_obj_.put ('aggregatesCount', aggs_rec_.count);
   aggs_obj_.put ('aggregateCols', aggs_list_);
END Initiate_Agg_Obj;

-----
-- Increment_Agg_Obj_From_Lf_Data()
-- Increment_Agg_Obj()
--   The aggregate object needs to be incremented as we read the data from our
--   data-source.  Therefore, the first function is called immediately when we
--   consume rows from the database, meaning that we only increment the record
--   count by one each time.  The second function works by combining aggregate
--   objects from sub-nodes.
--
PROCEDURE Increment_Agg_Obj_From_Lf_Data (
   aggs_obj_   IN OUT NOCOPY json_object_t,
   data_range_ IN     tp_cell_range,
   row_        IN     PLS_INTEGER )
IS
   aggs_list_    json_array_t := aggs_obj_.get_array('aggregateCols');
   agg_obj_      json_object_t;
   col_          PLS_INTEGER;
   rg_col_start_ PLS_INTEGER  := data_range_.tl.c - 1;
   rec_count_    PLS_INTEGER  := aggs_obj_.get_number('recordCount');
   value_        NUMBER;
BEGIN
   FOR ix_ IN 0 .. aggs_list_.get_size-1 LOOP
      agg_obj_ := treat (aggs_list_.get(ix_) as json_object_t);
      CASE agg_obj_.get_string('fn')
         WHEN 'count' THEN
            value_ := nvl (agg_obj_.get_number('value'), 0);
            agg_obj_.put ('value', value_ + 1);
         WHEN 'sum' THEN
            col_   := rg_col_start_ + agg_obj_.get_number('colId');
            value_ := nvl (agg_obj_.get_number('value'), 0);
            agg_obj_.put (
               'value', value_ + Get_Cell_Value_Num (col_, row_, data_range_.sheet_id)
            );
      END CASE;
      aggs_list_.put (ix_, agg_obj_, true);
   END LOOP;
   aggs_obj_.put ('recordCount', rec_count_ + 1);
   aggs_obj_.put ('aggregateCols', aggs_list_);
END Increment_Agg_Obj_From_Lf_Data;

PROCEDURE Increment_Agg_Obj (
   accum_aggs_obj_ IN OUT NOCOPY json_object_t,
   new_agg_values_ IN            json_object_t )
IS
   changed_        BOOLEAN      := false;
   accum_aggs_arr_ json_array_t := accum_aggs_obj_.get_array('aggregateCols');
   new_aggs_arr_   json_array_t := new_agg_values_.get_array('aggregateCols');
   accum_agg_      json_object_t;
   accum_val_      NUMBER;
   new_val_        NUMBER;
BEGIN
   FOR ix_ IN 0 .. accum_aggs_arr_.get_size-1 LOOP
      accum_agg_ := treat (accum_aggs_arr_.get(ix_) as json_object_t);
      IF accum_agg_.get_string('fn') IN ('count','sum') THEN
         changed_   := true;
         accum_val_ := accum_agg_.get_number('value');
         new_val_   := treat (new_aggs_arr_.get(ix_) as json_object_t).get_number('value');
         IF accum_val_ IS null AND new_val_ IS null THEN
            accum_agg_.put ('value', to_number(null));
         ELSE
            accum_agg_.put ('value', nvl(accum_val_,0) + nvl(new_val_,0));
         END IF;
      END IF;
   END LOOP;
   IF changed_ THEN
      accum_aggs_obj_.put ('aggregateCols', accum_aggs_arr_);
   END IF;
END Increment_Agg_Obj;

-----
-- Append_Array_Of_Len()
--   This function should be read as: "Append an extra arr_width_ items to the
--   end of the existing array", not "Make this current array to be arr_width_
--   items long"
--
PROCEDURE Append_Array_Of_Len (
   lv_arr_      IN OUT NOCOPY json_array_t,
   arr_width_   IN PLS_INTEGER,
   child_is_lf_ IN BOOLEAN,
   shared_item_ IN VARCHAR2,
   aggregates_  IN tp_col_agg_fns )
IS BEGIN
   IF child_is_lf_ THEN
      FOR i_ IN 1 .. arr_width_ LOOP
         lv_arr_.append (CASE i_ WHEN 1 THEN shared_item_ ELSE '' END);
      END LOOP;
   ELSE -- not child_is_lf_
      FOR i_ IN 1 .. arr_width_ - aggregates_.count LOOP
         lv_arr_.append (CASE i_ WHEN 1 THEN shared_item_ ELSE '' END);
      END LOOP;
      FOR ix_ IN 1 .. aggregates_.count LOOP
         lv_arr_.append (shared_item_ || ' ' || aggregates_(ix_).col_tot_name);
      END LOOP;
   END IF;
END Append_Array_Of_Len;


PROCEDURE Append_Value_Or_Null (
   arr_1d_  IN OUT NOCOPY json_array_t,
   agg_obj_ IN json_object_t )
IS
   agg_count_ PLS_INTEGER  := agg_obj_.get_number('aggregatesCount');
   aggs_arr_  json_array_t := agg_obj_.get_array('aggregateCols');
   rec_count_ PLS_INTEGER  := nvl (agg_obj_.get_number('recordCount'), 0);
BEGIN
   IF rec_count_ = 0 THEN
      arr_1d_ := json_array_t('[]');
   ELSE
      FOR ix_ IN 0 .. agg_count_-1 LOOP
         IF rec_count_ = 0 THEN
            arr_1d_.append(to_number(null));
         ELSE
            arr_1d_.append (
               treat (aggs_arr_.get(ix_) as json_object_t).get_number('value')
            );
         END IF;
      END LOOP;
   END IF;
END Append_Value_Or_Null;



-----
-- Build_Excel_Agg()
--   Builds out a 2D array representation of the pivot table as it will appear
--   in the final Excel grid.  There is one grid to represent the header cells
--   of the pivot table, and another to represent the data in the grid.  It is
--   best to think of this as being called on the leaf nodes (although it also
--   can start from the second node up in the case of the table-header).
--   This function is called (i.e. an aggregate object is created) only if the
--   number of matching records in the dataset is greater than zero.
--
PROCEDURE Build_Excel_Agg (
   aggs_obj_  IN OUT NOCOPY json_object_t )
IS
   xl_hd_lv_   json_array_t := json_array_t();
   xl_hd_full_ json_array_t := json_array_t();
BEGIN
   IF aggs_obj_.get_number('aggregatesCount') > 1 THEN
      FOR ix_ IN 0 .. aggs_obj_.get_array('aggregateCols').get_size-1 LOOP
         xl_hd_lv_.append (
            treat (
               aggs_obj_.get_array('aggregateCols').get(ix_) as json_object_t
            ).get_string('colName')
         );
      END LOOP;
      xl_hd_full_.append(xl_hd_lv_);
      aggs_obj_.put ('xlPtHead', xl_hd_full_);
   END IF;
END Build_Excel_Agg;

-----
-- Append_Arr_2D()
--   This function is called in a loop with lv0_arr_ starting life as an empty
--   array.  Remember that the array is 2-dimensional so each loop extends the
--   inner arrays.  The function works on both the header and main grid of our
--   pivot table
--     [[a, b, c],[d, e, f]] + [[g, h],[j, k]] => [[a, b, c, g, h][d, e, f, j, k]]
PROCEDURE Append_Arr_2D (
   lv0_arr_    IN OUT NOCOPY json_array_t, -- 2D
   lv1_arr_    IN json_array_t )           -- 2D
IS
   row_arr_ json_array_t;
BEGIN
   IF lv0_arr_.get_size = 0 THEN
      lv0_arr_ := lv1_arr_;
   ELSE
      FOR row_ IN 0 .. (lv1_arr_.get_size-1) LOOP -- for each "row" of the child array
         row_arr_ := treat (lv0_arr_.get(row_) as json_array_t);
         row_arr_.append_all (treat(lv1_arr_.get(row_) as json_array_t));
         lv0_arr_.put (row_, row_arr_, true);
      END LOOP;
   END IF;
END Append_Arr_2D;

-----
-- Append_Arr_Object_2D()
--   This, admittedly is a little messy.  But I'll try to demonstrate how it's
--   designed.  The two incoming arrays, and return array are all of the built
--   in the following format:
--      [
--        {"lv":4, "grid":[null, 56, null]}, {"lv":4, "grid":[46, 12, null]}, ...
--      ]
--   The function is called from a loop which cumulatively appends lv1_arr_ to
--   lv0_arr_ on an iteration.  If the loop is building in the horizontal, the
--   outermost array will have only one object element, and the "grid" node of
--   that object will be extended with the corresponding node of lv1_arr_.  If
--   we build vertically, we simply append the arrays together.
--
PROCEDURE Append_Arr_Object_2D (
   lv0_arr_    IN OUT NOCOPY json_array_t, -- array of object, as per description above
   lv1_arr_    IN json_array_t,            -- array of object
   horizontal_ IN BOOLEAN,
   level_      IN PLS_INTEGER )
IS
   new_grid_arr_ json_array_t;
   rtn_grid_arr_ json_array_t  := json_array_t();
   lv_obj_       json_object_t := json_object_t();
BEGIN
   IF horizontal_ AND lv1_arr_.get_size > 1 THEN
      Raise_App_Error ('Array should only be of length 1 when building in the horizontal direction');
   END IF;
   IF horizontal_ THEN
      IF lv0_arr_.get_size > 0 THEN
         rtn_grid_arr_ := treat(lv0_arr_.get(0) as json_object_t).get_array('grid');
      END IF;
      new_grid_arr_ := treat (lv1_arr_.get(0) as json_object_t).get_array('grid');
      rtn_grid_arr_.append_all(new_grid_arr_);
      lv_obj_.put ('lv', level_);
      lv_obj_.put ('grid', rtn_grid_arr_);
      IF lv0_arr_.get_size = 0 THEN
         lv0_arr_.append (lv_obj_);
      ELSE
         lv0_arr_.put (0, lv_obj_, true);
      END IF;
   ELSE
      lv0_arr_.append_all (lv1_arr_);
   END IF;
END Append_Arr_Object_2D;

-----
-- Append_Aggs_Build_Next_2D()
-- Finish_Grid_Group()
--   These functions use the same array structures as described above, but are
--   called at the end of the loop to calculate additional aggregates.
--
PROCEDURE Append_Aggs_Build_Next_2D (
   lv0_arr_1d_ IN OUT NOCOPY json_array_t,
   lv1_arr_2d_ IN OUT NOCOPY json_array_t,
   h_level_    IN PLS_INTEGER,
   aggregates_ IN tp_col_agg_fns )
IS
   lv_arr_   json_array_t;
   col_name_ VARCHAR2(32000);
BEGIN
   FOR ag_ IN 1 .. aggregates_.count LOOP
      col_name_ := CASE
         WHEN aggregates_.count = 1 THEN 'Grand Total'
         ELSE aggregates_(ag_).col_tot_name
      END;
      lv0_arr_1d_.append (CASE h_level_ WHEN 1 THEN col_name_ ELSE '' END);
      -- We can't `put()` an element into location-zero of an empty array, and
      -- so this IF statement becomes necessary
      IF lv1_arr_2d_.get_size > 0 THEN
         FOR lv_ IN 0 .. lv1_arr_2d_.get_size - 1 LOOP
            lv_arr_ := treat (lv1_arr_2d_.get(lv_) as json_array_t);
            lv_arr_.append ('');
            lv1_arr_2d_.put (lv_, lv_arr_, true);
         END LOOP;
         lv1_arr_2d_.put (0, lv0_arr_1d_);
      ELSE
         lv1_arr_2d_.append (lv0_arr_1d_);
      END IF;
   END LOOP;
END Append_Aggs_Build_Next_2D;

PROCEDURE Finish_Grid_Group (
   pt_group_obj_arr_ IN OUT NOCOPY json_array_t,
   node_agg_obj_     IN json_object_t,
   build_horizontal_ IN BOOLEAN,
   level_            IN PLS_INTEGER )
IS

   agg_count_ PLS_INTEGER   := node_agg_obj_.get_number('aggregatesCount');
   accum_arr_ json_array_t  := json_array_t();
   accum_obj_ json_object_t := json_object_t();
   agg_obj_   json_object_t;
   row_obj_   json_object_t;
   row_arr_   json_array_t;
   val_       NUMBER;

   PROCEDURE Init_Accum_Arr (
      target_size_ IN PLS_INTEGER )
   IS BEGIN
      IF accum_arr_.get_size = 0 THEN
         FOR i_ IN 1 .. target_size_ LOOP
            accum_arr_.append(to_number(null));
         END LOOP;
      END IF;
   END Init_Accum_Arr;

BEGIN
   IF build_horizontal_ THEN
      row_obj_ := treat (pt_group_obj_arr_.get(0) as json_object_t);
      row_arr_ := row_obj_.get_array('grid');
      FOR ix_ IN 0 .. agg_count_-1 LOOP
         agg_obj_ := treat (node_agg_obj_.get_array('aggregateCols').get(ix_) as json_object_t);
         row_arr_.append (agg_obj_.get_number('value'));
      END LOOP;
      row_obj_.put ('lv', level_);
      row_obj_.put ('grid', row_arr_);
      pt_group_obj_arr_.put (0, row_obj_, true);

   ELSE
      Init_Accum_Arr (treat(pt_group_obj_arr_.get(0) as json_object_t).get_array('grid').get_size);
      FOR r_ IN 0 .. pt_group_obj_arr_.get_size-1 LOOP
         row_obj_ := treat (pt_group_obj_arr_.get(r_) as json_object_t);
         IF level_+1 = row_obj_.get_number('lv') THEN
            row_arr_ := row_obj_.get_array('grid');
            FOR c_ IN 0 .. row_arr_.get_size-1 LOOP
               IF row_arr_.get_number(c_) IS NOT null THEN
                  val_ := row_arr_.get_number(c_) + nvl(accum_arr_.get_number(c_),0);
                  accum_arr_.put (c_, val_, true);
               END IF;
            END LOOP;
         END IF;
      END LOOP;
      accum_obj_.put ('lv', level_);
      accum_obj_.put ('grid', accum_arr_);
      IF level_ != 1 THEN
         pt_group_obj_arr_.put (0, accum_obj_);
      ELSE
         pt_group_obj_arr_.append (accum_obj_);
      END IF;
   END IF;
END Finish_Grid_Group;

-----
-- Complete_Pivot_Header()
--
--
PROCEDURE Complete_Pivot_Header (
   xl_pt_hd_arr_ IN OUT NOCOPY json_array_t,
   aggregates_   IN tp_col_agg_fns )
IS
   th_width_ PLS_INTEGER;
   th_depth_ PLS_INTEGER;
   lv_arr_   json_array_t := json_array_t();
BEGIN
   IF xl_pt_hd_arr_.get_size = 0 THEN
      FOR col_ IN 0 .. aggregates_.count LOOP
         lv_arr_.append (CASE
            WHEN col_ = 0 THEN 'Column Labels'
            ELSE aggregates_(col_).col_tot_name
         END);
      END LOOP;
      xl_pt_hd_arr_.append(lv_arr_);
   ELSE
      th_width_ := treat (xl_pt_hd_arr_.get(0) as json_array_t).get_size;
      FOR ix_ IN 1 .. th_width_ LOOP
         lv_arr_.append(CASE ix_ WHEN 1 THEN 'Column Labels' ELSE '' END);
      END LOOP;
      xl_pt_hd_arr_.put (0, lv_arr_, false);
      th_depth_ := xl_pt_hd_arr_.get_size - 1;
      FOR lv_ IN 0 .. th_depth_ LOOP
         lv_arr_ := treat (xl_pt_hd_arr_.get(lv_) as json_array_t);
         lv_arr_.put (0, CASE
            WHEN lv_ = th_depth_                   THEN 'Row Labels'
            WHEN lv_ = 0 AND aggregates_.count = 1 THEN aggregates_(1).col_tot_name
            ELSE ''
         END);
         xl_pt_hd_arr_.put (lv_, lv_arr_, true);
      END LOOP;
   END IF;
END Complete_Pivot_Header;

-----
-- Pivot_Header_To_Cols()
--   Converts the pivot-table's header grid (which has already been calculated
--   as a 2D array) to a vertical format, helping us build the <colItems> node
--   of the pivotTable part.  Remember that the function above hasn't yet been
--   called, and so "finishing" has been applied yet.
--     [["Fruit", "", "Veg", "","Total"]["Apple","Pear","Carrot","Sprout",""]
--       => [["Fruit","Apple"],["","Pear"],["Veg","Carrot"],["","Sprout"],["Total",""]]
--
FUNCTION Pivot_Header_To_Cols (
   xl_pt_hd_arr_ IN json_array_t,
   ds_range_     IN tp_cell_range,
   h_rollups_    IN tp_pivot_cols ) RETURN json_object_t
IS
   ix_           PLS_INTEGER;
   th_width_     PLS_INTEGER   := treat (xl_pt_hd_arr_.get(0) as json_array_t).get_size;
   col_name_arr_ json_array_t  := json_array_t();
   int_arr_      json_array_t;
   rtn_arr_      json_array_t  := json_array_t();
   rtn_obj_      json_object_t := json_object_t();
BEGIN
   ix_ := h_rollups_.first;
   WHILE ix_ IS NOT null LOOP
      col_name_arr_.append (
         Range_Col_Head_Name (ds_range_, col_offs_ => h_rollups_(ix_))
      );
      ix_  := h_rollups_.next(ix_);
   END LOOP;
   FOR int_loop_ IN 0 .. th_width_-1 LOOP
      int_arr_ := json_array_t();
      FOR ext_loop_ IN 0 .. xl_pt_hd_arr_.get_size-1 LOOP
         int_arr_.append (
            treat (xl_pt_hd_arr_.get(ext_loop_) as json_array_t).get_string(int_loop_)
         );
      END LOOP;
      rtn_arr_.append (int_arr_);
   END LOOP;
   rtn_obj_.put ('colNames', col_name_arr_);
   rtn_obj_.put ('vHead',    rtn_arr_);
   RETURN rtn_obj_;
END Pivot_Header_To_Cols;

-----
-- Breadcrumb_Is_In_Axes()
--   To reduce computation, this function looks at the axes of our pivot table
--   to see if the combination of X and Y axes exists.  In some cases the axes
--   might exist even if there's no record in the "database" which matches our
--   breadcrumb trail.  However, when the XY axes does not exist we can notify
--   the caller that no further calculation is required.
--
FUNCTION Breadcrumb_Is_In_Axes (
   breadcrumb_ VARCHAR2,
   v_axes_     json_object_t,
   h_axes_     json_object_t,
   v_depth_    PLS_INTEGER ) RETURN BOOLEAN
IS
   pos_      PLS_INTEGER   := instr (breadcrumb_, '/', 1, (v_depth_+1));
   v_trail_  VARCHAR2(200) := substr (breadcrumb_, 1, pos_-1);
   h_trail_  VARCHAR2(200) := substr (breadcrumb_, pos_);
   in_axes_  BOOLEAN       := true;
   axes_obj_ json_object_t;
   trail_    VARCHAR2(200);

   CURSOR get_shared_items IS
      SELECT regexp_substr (trail_, '[^/]+', 1, level) si_name, level lv
      FROM   dual
      CONNECT BY regexp_substr (trail_, '[^/]+', 1, level) IS NOT null
      ORDER BY lv ASC;

BEGIN
   IF pos_ = 0 THEN
      v_trail_ := breadcrumb_;
      h_trail_ := '';
   END IF;
   FOR i_ IN 1 .. 2 LOOP -- 1 = vertical, 2 = horizontal
      trail_    := CASE i_ WHEN 1 THEN v_trail_ ELSE h_trail_ END;
      axes_obj_ := CASE i_ WHEN 1 THEN v_axes_  ELSE h_axes_  END;
      IF trail_ IS NOT null THEN
         FOR si_ IN get_shared_items LOOP
            in_axes_ := axes_obj_.get_object('sharedItems').has(si_.si_name);
            EXIT WHEN not in_axes_;
            axes_obj_ := axes_obj_.get_object('sharedItems').get_object(si_.si_name);
         END LOOP;
      END IF;
      EXIT WHEN not in_axes_;
   END LOOP;
   RETURN in_axes_;
END Breadcrumb_Is_In_Axes;


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
   v_level_       IN PLS_INTEGER    := 0, -- should be length(h_filter_vals_); [(2) => "val1", (3) => "val2"]
   h_level_       IN PLS_INTEGER    := 0,
   v_filter_vals_ IN tp_col_filters := tp_col_filters(),
   h_filter_vals_ IN tp_col_filters := tp_col_filters(),
   extra_filters_ IN tp_col_filters := tp_col_filters(),
   direction_     IN VARCHAR2       := 'top',
   breadcrumb_    IN VARCHAR2       := '',
   v_axes_        IN json_object_t  := json_object_t(),
   h_axes_        IN json_object_t  := json_object_t() ) RETURN json_object_t
IS

   dir_horizontal_   CONSTANT BOOLEAN := direction_ = 'horizontal';
   dir_vertical_     CONSTANT BOOLEAN := direction_ = 'vertical';
   dir_full_grid_    CONSTANT BOOLEAN := direction_ = 'full-grid';

   pt_               tp_pivot_table := wb_.pivot_tables(pivot_id_);
   cache_            tp_pivot_cache := wb_.pivot_caches(pt_.cache_id);
   ds_range_         tp_cell_range  := wb_.pivot_caches(pt_.cache_id).ds_range;
   aggregates_       tp_col_agg_fns := pt_.pivot_axes.col_agg_fns; -- (1) => tp_agg_fn(colid:3, agg_fn:sum);
   child_agg_obj_    json_object_t;
   sis_obj_          json_object_t  := json_object_t(); -- sis: shared items
   results_obj_      json_object_t  := json_object_t();
   child_obj_        json_object_t  := json_object_t();
   xl_lf_obj_        json_object_t  := json_object_t();
   lf_agg_obj_       json_object_t;
   lf_arr_           json_array_t   := json_array_t();
   xl_lf_arr_        json_array_t   := json_array_t();
   lv0_1d_hd_arr_    json_array_t   := json_array_t();
   lv0_2d_hd_arr_    json_array_t   := json_array_t();
   lv0_grid_obj_arr_ json_array_t   := json_array_t();
   axes_obj_         json_object_t  := json_object_t();
   axes_arr_         json_array_t   := json_array_t();
   xl_pt_axes_arr_   json_array_t   := json_array_t();
   xl_pt_hd_v_obj_   json_object_t;
   node_agg_obj_     json_object_t;

   next_v_level_     PLS_INTEGER;
   next_h_level_     PLS_INTEGER;
   v_depth_          PLS_INTEGER    := pt_.pivot_axes.vrollups.count;
   h_depth_          PLS_INTEGER    := pt_.pivot_axes.hrollups.count; -- [1 => 3, 2 => 4, 3 => 1], key is always sequential
   v_filters_        tp_col_filters := v_filter_vals_;
   h_filters_        tp_col_filters := h_filter_vals_; -- order not important, only a filter
   comb_filters_     tp_col_filters;

   is_h_leaf_        BOOLEAN        := h_level_ = h_depth_;
   goto_leaf_        BOOLEAN        := h_level_ > h_depth_ OR (h_level_ > 0 AND dir_vertical_);
   last_v_roll_      BOOLEAN        := v_level_ = v_depth_ AND h_level_ = 0;
   roll_vertical_    BOOLEAN        := not goto_leaf_
      AND (last_v_roll_ OR v_level_ < v_depth_) AND (dir_vertical_ OR dir_full_grid_);
   roll_horizntl_    BOOLEAN        := not goto_leaf_ AND (dir_horizontal_ OR dir_full_grid_);
   unroll_horzntl_   BOOLEAN        := h_level_ > 0;
   leaf_h_level_     PLS_INTEGER    := CASE WHEN goto_leaf_ THEN h_level_-1 ELSE h_level_ END;

   direct_nodes_     PLS_INTEGER    := 0;
   record_count_     PLS_INTEGER    := 0;
   lv_width_         PLS_INTEGER    := 0;
   lv_height_        PLS_INTEGER    := 0;
   col_id_           PLS_INTEGER;
   child_width_      PLS_INTEGER;
   child_is_lf_      BOOLEAN;
   col_name_         VARCHAR2(32000);
   shared_item_      VARCHAR2(32000);
   fc_offset_        PLS_INTEGER;   -- fc:filter-column
   rec_count_        PLS_INTEGER    := 0;
   grid_nd_exists_   BOOLEAN        := false;
   calc_required_    BOOLEAN;
   keep_row_         BOOLEAN;

BEGIN

   IF leaf_h_level_ > h_depth_ OR v_level_ > v_depth_ THEN
      Raise_App_Error ('h-level / v-level are: :P1 / :P2', to_char(leaf_h_level_), to_char(v_level_));
   END IF;

   -- ***
   -- *** Root level initiator; This is also where the recursion exits
   IF direction_ = 'top' THEN
      results_obj_.put ('v-tree', CASE
         WHEN v_depth_ = 0 THEN json_object_t()
         ELSE Json_Aggregates_From_Filters (
            pivot_id_, v_level_ => 1, h_level_ => 0, extra_filters_ => extra_filters_, direction_ => 'vertical'
         )
      END);
      results_obj_.put ('h-tree', CASE
         WHEN h_depth_ = 0 THEN json_object_t()
         ELSE Json_Aggregates_From_Filters (
            pivot_id_, v_level_ => 0, h_level_ => 1, extra_filters_ => extra_filters_, direction_ => 'horizontal'
         )
      END);
      results_obj_.put ('full-grid', CASE
         WHEN v_depth_ = 0 THEN json_object_t()
         ELSE Json_Aggregates_From_Filters (
            pivot_id_, v_level_ => 1, h_level_ => 0, extra_filters_ => extra_filters_, direction_ => 'full-grid',
            v_axes_ => results_obj_.get_object('v-tree'), h_axes_ => results_obj_.get_object('h-tree')
         )
      END);

      -- complete the vertical axes
      IF results_obj_.get_object('v-tree').has('aggregates') THEN
         child_agg_obj_ := results_obj_.get_object('v-tree').get_object('aggregates');
         IF child_agg_obj_.has('xlPtvAxes') THEN
            xl_pt_axes_arr_ := child_agg_obj_.get_array('xlPtvAxes');
            axes_obj_.put ('lv', 0);
            axes_obj_.put ('val', 'Grand Total');
            xl_pt_axes_arr_.append (axes_obj_);
            results_obj_.put ('xlPtvAxes', xl_pt_axes_arr_);
         END IF;
      END IF;

      -- complete the horizontal axes
      IF results_obj_.get_object('h-tree').has('aggregates') THEN
         child_agg_obj_ := results_obj_.get_object('h-tree').get_object('aggregates');
         IF child_agg_obj_.has('xlPtHead') THEN
            xl_pt_axes_arr_ := results_obj_.get_object('h-tree').get_object('aggregates').get_array('xlPtHead');
            xl_pt_hd_v_obj_ := Pivot_Header_To_Cols (xl_pt_axes_arr_, ds_range_, pt_.pivot_axes.hrollups);
            Complete_Pivot_Header (xl_pt_axes_arr_, aggregates_);
            results_obj_.put ('xlPthAxes', xl_pt_hd_v_obj_);
            results_obj_.put ('xlPtHead', xl_pt_axes_arr_);
         END IF;
      END IF;

      -- Now finish the job by putting that json into an Excel sheet!
      Unravel_Json_To_Sheet (pivot_id_, results_obj_);


   -- ***
   -- *** Build a mid-level node
   ELSIF roll_vertical_ OR roll_horizntl_ THEN

      IF roll_vertical_ THEN
         col_id_       := pt_.pivot_axes.vrollups(v_level_);
         next_v_level_ := v_level_ + CASE WHEN last_v_roll_ THEN 0 ELSE 1 END;
         next_h_level_ := h_level_ + CASE WHEN last_v_roll_ THEN 1 ELSE 0 END;
      ELSE
         col_id_       := pt_.pivot_axes.hrollups(h_level_);
         next_v_level_ := v_level_;
         next_h_level_ := h_level_ + 1;
      END IF;

      col_name_ := Range_Col_Head_Name (cache_.ds_range, col_id_);
      Initiate_Agg_Obj (pt_.pivot_axes.col_agg_fns, node_agg_obj_);

      shared_item_ := cache_.cached_fields(col_name_).shared_items.first;
      WHILE shared_item_ IS NOT null LOOP

         IF roll_vertical_ THEN
            v_filters_(col_id_) := shared_item_;
         ELSE
            h_filters_(col_id_) := shared_item_;
         END IF;
         child_obj_ := Json_Aggregates_From_Filters (
            pivot_id_      => pivot_id_,
            v_level_       => next_v_level_,
            h_level_       => next_h_level_,
            v_filter_vals_ => v_filters_,
            h_filter_vals_ => h_filters_,
            extra_filters_ => extra_filters_,
            direction_     => direction_,
            breadcrumb_    => breadcrumb_ || '/' || shared_item_,
            v_axes_        => v_axes_,
            h_axes_        => h_axes_
         );
         IF dir_full_grid_ OR child_obj_.get_number('recordCount') > 0 THEN

            child_agg_obj_ := child_obj_.get_object('aggregates');
            sis_obj_.put (shared_item_, child_obj_);
            Increment_Agg_Obj (node_agg_obj_, child_agg_obj_);
            direct_nodes_ := direct_nodes_ + 1;
            record_count_ := record_count_ + child_obj_.get_number('recordCount');

            IF dir_horizontal_ THEN
               child_width_ := child_obj_.get_number('width');
               child_is_lf_ := child_obj_.get_boolean('isLeaf');
               lv_width_    := lv_width_ + child_width_;
               Append_Array_Of_Len (lv0_1d_hd_arr_, child_width_, child_is_lf_, shared_item_, aggregates_);
               -- Leaf only includes "xlPtHead" if there's more than one aggregate count
               IF child_agg_obj_.has('xlPtHead') THEN
                  Append_Arr_2D (lv0_2d_hd_arr_, child_agg_obj_.get_array('xlPtHead'));
               END IF;
            ELSIF dir_vertical_ THEN
               lv_height_ := lv_height_ + child_obj_.get_number('height');
               axes_obj_.put ('lv', v_level_);
               axes_obj_.put ('colName', col_name_);
               axes_obj_.put ('val', shared_item_);
               axes_arr_.append (axes_obj_);
               IF child_agg_obj_.has ('xlPtvAxes') THEN
                  axes_arr_.append_all (child_agg_obj_.get_array ('xlPtvAxes'));
               END IF;
            ELSIF dir_full_grid_ THEN
               IF child_obj_.has('xlGrid') THEN
                  grid_nd_exists_ := true;
                  Append_Arr_Object_2D (
                     lv0_grid_obj_arr_, child_obj_.get_array('xlGrid'),
                     unroll_horzntl_, next_v_level_+next_h_level_
                  );
               END IF;
            END IF;

         END IF;

         shared_item_ := cache_.cached_fields(col_name_).shared_items.next(shared_item_);
      END LOOP;

      IF dir_horizontal_ THEN
         Append_Aggs_Build_Next_2D (lv0_1d_hd_arr_, lv0_2d_hd_arr_, h_level_, aggregates_);
         lv_width_ := lv_width_ + node_agg_obj_.get_number('aggregatesCount');
         Build_Excel_Agg (node_agg_obj_);
         node_agg_obj_.put ('xlPtHead', lv0_2d_hd_arr_);
      ELSIF dir_vertical_ THEN
         lv_height_ := lv_height_ + 1;
         node_agg_obj_.put ('xlPtvAxes', axes_arr_);
      ELSIF dir_full_grid_ THEN
         IF grid_nd_exists_ THEN
            Finish_Grid_Group (
               lv0_grid_obj_arr_, node_agg_obj_, unroll_horzntl_, h_level_+v_level_
            );
         END IF;
      END IF;

      results_obj_.put ('direction',       CASE WHEN roll_vertical_ THEN 'vertical' ELSE 'horizontal' END);
      results_obj_.put ('colId',           col_id_);
      results_obj_.put ('colDesc',         col_name_);
      results_obj_.put ('breadcrumb',      breadcrumb_);
      results_obj_.put ('isLeaf',          false);
      results_obj_.put ('sharedItemCount', direct_nodes_);
      results_obj_.put ('recordCount',     record_count_);
      IF dir_horizontal_ THEN
         results_obj_.put ('width',        lv_width_);
         results_obj_.put ('level',        h_level_);
      ELSIF dir_vertical_ THEN
         results_obj_.put ('height',       lv_height_);
         results_obj_.put ('level',        v_level_);
      ELSIF dir_full_grid_ THEN
         results_obj_.put ('h-level',      h_level_);
         results_obj_.put ('v-level',      least (v_level_,v_depth_));
         results_obj_.put ('is-hLeaf',     is_h_leaf_);
         results_obj_.put ('is-vLeaf',     v_depth_<=v_level_);
      END IF;
      results_obj_.put ('sharedItems',     sis_obj_);
      results_obj_.put ('aggregates',      node_agg_obj_);
      IF grid_nd_exists_ THEN
         results_obj_.put ('xlGrid', lv0_grid_obj_arr_);
      END IF;


   -- ***
   -- *** Build the leaf node
   ELSE

      Initiate_Agg_Obj (aggregates_, node_agg_obj_);
      comb_filters_  := Combine_Arrays (v_filter_vals_, h_filter_vals_);
      calc_required_ := dir_vertical_ OR dir_horizontal_ OR (
         dir_full_grid_ AND Breadcrumb_Is_In_Axes (breadcrumb_, v_axes_, h_axes_, v_depth_)
      );

      -- At the leaf of our Json grid we interogate the "database" (built into
      -- the Excel sheet) to sum up aggregates from source.  Parent nodes base
      -- their own calculations on these child calculations (as opposed to re-
      -- querying the same database)
      --
      IF calc_required_ THEN
         FOR r_ IN (ds_range_.tl.r+1) .. ds_range_.br.r LOOP -- loop on dataset rows
            keep_row_  := true;
            fc_offset_ := comb_filters_.first;
            WHILE fc_offset_ IS NOT null AND keep_row_ LOOP -- loop on search criteria
               col_id_    := (ds_range_.tl.c-1) + fc_offset_;
               keep_row_  := keep_row_ AND Get_Cell_Value_Raw (col_id_, r_, ds_range_.sheet_id, false) = comb_filters_(fc_offset_);
               fc_offset_ := comb_filters_.next(fc_offset_);
            END LOOP;
            IF keep_row_ THEN
               rec_count_ := rec_count_ + 1;
               Increment_Agg_Obj_From_Lf_Data (node_agg_obj_, ds_range_, r_);
            END IF;
         END LOOP;
         FOR ix_ IN 0 .. aggregates_.count-1 LOOP
            IF rec_count_ = 0 THEN
               lf_arr_.append (to_number(null));
            ELSE
               lf_agg_obj_ := treat (node_agg_obj_.get_array('aggregateCols').get(ix_) as json_object_t);
               lf_arr_.append (lf_agg_obj_.get_number('value'));
            END IF;
         END LOOP;
         xl_lf_obj_.put ('lv', v_level_ + h_level_);
         xl_lf_obj_.put ('grid', lf_arr_);
         xl_lf_arr_.append (xl_lf_obj_);
      END IF;

      results_obj_.put ('breadcrumb',  breadcrumb_);
      results_obj_.put ('isLeaf',      true);
      results_obj_.put ('recordCount', rec_count_);
      results_obj_.put ('width',       aggregates_.count);
      results_obj_.put ('height',      1);
      results_obj_.put ('aggregates',  node_agg_obj_);
      IF dir_full_grid_ AND calc_required_ THEN
         results_obj_.put ('xlGrid',      xl_lf_arr_);
      END IF;

   END IF;

   RETURN results_obj_;

END Json_Aggregates_From_Filters;


-----
-- Finish_Pivot_Tables()
--   Must be called after Build_Pivot_Caches_And_Tables(), which isn't hard to
--   achieve on account of being called at the very beginning of the procedure
--   Finish().  Worth noting anyway, to avoid potential problem in future.
--
PROCEDURE Finish_Pivot_Tables (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_         dbms_XmlDom.DomDocument;
   attrs_       nyce_xml.xml_attrs_arr;
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
   pt_axes_     tp_pivot_axes;
   shared_item_ VARCHAR2(32000);
   agg_col_     PLS_INTEGER;

BEGIN

   FOR pt_ IN 1 .. wb_.pivot_tables.count LOOP

      j_piv_ := Json_Aggregates_From_Filters (pivot_id_ => pt_);

      cache_   := wb_.pivot_caches(wb_.pivot_tables(pt_).cache_id);
      pt_axes_ := wb_.pivot_tables(pt_).pivot_axes;

      -- xl/pivotTables/pivotTable:P1.xml
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
      nyce_xml.attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
      nyce_xml.attr ('mc:Ignorable', 'xr', attrs_);
      nyce_xml.attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
      nyce_xml.attr ('xr:uid', Get_Guid, attrs_);
      nyce_xml.attr ('name', wb_.pivot_tables(pt_).pivot_name, attrs_);
      nyce_xml.attr ('cacheId', wb_.pivot_tables(pt_).cache_id, attrs_);
      nyce_xml.attr ('applyNumberFormats', '0', attrs_);
      nyce_xml.attr ('applyBorderFormats', '0', attrs_);
      nyce_xml.attr ('applyFontFormats', '0', attrs_);
      nyce_xml.attr ('applyPatternFormats', '0', attrs_);
      nyce_xml.attr ('applyAlignmentFormats', '0', attrs_);
      nyce_xml.attr ('applyWidthHeightFormats', '1', attrs_);
      nyce_xml.attr ('dataCaption', 'Values', attrs_);
      nyce_xml.attr ('createdVersion', '7', attrs_); -- Version of Excel in which this pivot was created!
      nyce_xml.attr ('updatedVersion', '7', attrs_);
      nyce_xml.attr ('minRefreshableVersion', '3', attrs_); -- Minimum version of Excel which is compatible (apparently)
      nyce_xml.attr ('useAutoFormatting', '1', attrs_);
      nyce_xml.attr ('itemPrintTitles', '1', attrs_);
      nyce_xml.attr ('indent', '0', attrs_);
      nyce_xml.attr ('outline', '1', attrs_);
      nyce_xml.attr ('outlineData', '1', attrs_);
      nyce_xml.attr ('multipleFieldFilters', '0', attrs_);
      nd_ptd_ := Nyce_Xml.Make_Root_Node (doc_, 'pivotTableDefinition', attrs_);

      wb_.pivot_tables(pt_).pivot_height := j_piv_.get_object('v-tree').get_number('height') + j_piv_.get_array('xlPtHead').get_size;
      wb_.pivot_tables(pt_).pivot_width  := treat(j_piv_.get_array('xlPtHead').get(0) as json_array_t).get_size;
      pt_region_ := tp_cell_range (
         sheet_id => wb_.pivot_tables(pt_).on_sheet,
         tl       => wb_.pivot_tables(pt_).location_tl,
         br       => tp_cell_loc (
            c => wb_.pivot_tables(pt_).location_tl.c + wb_.pivot_tables(pt_).pivot_width - 1,
            r => wb_.pivot_tables(pt_).location_tl.r + wb_.pivot_tables(pt_).pivot_height - 1
         )
      );

      nyce_xml.natr ('ref', Alfan_Range(pt_region_), attrs_);
      nyce_xml.attr ('firstHeaderRow', '1', attrs_);
      nyce_xml.attr ('firstDataRow',   j_piv_.get_array('xlPtHead').get_size, attrs_);
      nyce_xml.attr ('firstDataCol',   '1', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'location', attrs_);

      -- <pivotFields>
      nyce_xml.natr ('count', to_char(cache_.cf_order.count), attrs_);
      nd_pfs_ := Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'pivotFields', attrs_);
      FOR cf_ix_ IN cache_.cf_order.first .. cache_.cf_order.last LOOP

         -- the rollup in the pivot-table needn't be the same as the rollup in
         -- the cache because a cache could serve multiple tables.  The rollup
         -- calculated here must therefore be that taken from the table
         nyce_xml.catr (attrs_);
         CASE Get_Agg_Fn_From_Axes (pt_axes_, cf_ix_)
            WHEN 'row' THEN nyce_xml.attr ('axis', 'axisRow', attrs_);
            WHEN 'col' THEN nyce_xml.attr ('axis', 'axisCol', attrs_);
            WHEN 'sum' THEN nyce_xml.attr ('dataField', '1', attrs_);
            -- count needs to be here too, probably
            ELSE null;
         END CASE;
         nyce_xml.attr ('showAll', 0, attrs_);
         nd_pf_ := Nyce_Xml.Xml_Node (doc_, nd_pfs_, 'pivotField', attrs_);

         cf_ := cache_.cached_fields(cache_.cf_order(cf_ix_));
         IF cf_.shared_items.count > 0 THEN
            nyce_xml.natr ('count', to_char(cf_.shared_items.count + 1), attrs_);
            nd_is_ := Nyce_Xml.Xml_Node (doc_, nd_pf_, 'items', attrs_);

            shared_item_ := cf_.shared_items.first;
            WHILE shared_item_ IS NOT null LOOP
               nyce_xml.natr ('x', cf_.shared_items(shared_item_), attrs_);
               Nyce_Xml.Xml_Node (doc_, nd_is_, 'item', attrs_);
               shared_item_ := cf_.shared_items.next(shared_item_);
            END LOOP;
            nyce_xml.natr ('t', 'default', attrs_);
            Nyce_Xml.Xml_Node (doc_, nd_is_, 'item', attrs_);
         END IF;

      END LOOP;

      -- cache row items (vertical) <rowFields> and <rowItems>
      IF pt_axes_.vrollups.count > 0 THEN
         nyce_xml.natr ('count', to_char(pt_axes_.vrollups.count), attrs_);
         nd_pfs_ := Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'rowFields', attrs_);
         FOR r_ IN 1 .. pt_axes_.vrollups.count LOOP
            nyce_xml.natr ('x', to_char(pt_axes_.vrollups(r_) - 1), attrs_);
            Nyce_Xml.Xml_Node (doc_, nd_pfs_, 'field', attrs_);
         END LOOP;
         nyce_xml.natr ('count', to_char(j_piv_.get_object('v-tree').get_number('height')), attrs_);
         nd_ri_ := Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'rowItems', attrs_);
         Unravel_Json_Ptv_Axes_Xml (doc_, nd_ri_, j_piv_.get_array('xlPtvAxes'), cache_);
      ELSE
         nyce_xml.natr ('count', '1', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'rowItems/i');
      END IF;

      -- cache col items (horizontal) <colFields> and <colItems>
      IF pt_axes_.hrollups.count > 0 THEN
         nyce_xml.natr ('count', to_char(pt_axes_.hrollups.count), attrs_);
         nd_pfs_ := Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'colFields', attrs_);
         FOR r_ IN 1 .. pt_axes_.hrollups.count LOOP
            nyce_xml.natr ('x', to_char(pt_axes_.hrollups(r_) - 1), attrs_);
            Nyce_Xml.Xml_Node (doc_, nd_pfs_, 'field', attrs_);
         END LOOP;
         nyce_xml.natr ('count', to_char(j_piv_.get_object('h-tree').get_number('width')), attrs_);
         nd_ri_ := Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'colItems', attrs_);
         Unravel_Json_Pth_Axes_Xml (doc_, nd_ri_, j_piv_.get_object('xlPthAxes'), cache_);
      ELSE
         nyce_xml.natr ('count', '1', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'colItems/i');
      END IF;

      -- <dataFields> are the aggregate columns
      nyce_xml.natr ('count', to_char(pt_axes_.col_agg_fns.count), attrs_);
      nd_dfs_ := Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'dataFields', attrs_);
      FOR ix_ IN 1 .. pt_axes_.col_agg_fns.count LOOP
         agg_col_ := pt_axes_.col_agg_fns(ix_).colid;
         nyce_xml.natr ('name', pt_axes_.col_agg_fns(ix_).col_agg_name, attrs_);
         nyce_xml.attr ('fld', agg_col_ - 1, attrs_); -- zero based, I think
         nyce_xml.attr ('baseField', '0', attrs_); -- used with showDataAs, which we aren't using for now
         nyce_xml.attr ('baseItem', '0', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_dfs_, 'dataField', attrs_);
      END LOOP;

      nyce_xml.natr ('name', 'PivotStyleLight16', attrs_);
      nyce_xml.attr ('showRowHeaders', '1', attrs_);
      nyce_xml.attr ('showColHeaders', '1', attrs_);
      nyce_xml.attr ('showRowStripes', '0', attrs_);
      nyce_xml.attr ('showColStripes', '0', attrs_);
      nyce_xml.attr ('showLastColumn', '1', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'pivotTableStyleInfo', attrs_);

      nd_exl_ := Nyce_Xml.Xml_Node (doc_, nd_ptd_, 'extLst');

      nyce_xml.natr ('uri', Get_Guid, attrs_);
      nyce_xml.attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
      nd_ext_ := Nyce_Xml.Xml_Node (doc_, nd_exl_, 'ext', attrs_);

      nyce_xml.natr ('hideValuesRow', '1', attrs_);
      nyce_xml.attr ('xmlns:xm', 'http://schemas.microsoft.com/office/excel/2006/main', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ext_, 'pivotTableDefinition', 'x14', attrs_);

      nyce_xml.natr ('uri', Get_Guid, attrs_);
      nyce_xml.attr ('xmlns:xpdl', 'http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout', attrs_);
      nd_ext_ := Nyce_Xml.Xml_Node (doc_, nd_exl_, 'ext', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ext_, 'pivotTableDefinition16', 'xpdl');

      Add1Xml (excel_, rep('xl/pivotTables/pivotTable:P1.xml',pt_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);


      -- One _rel file per pivot table.
      -- xl/pivotTables/_rels/pivotTable:P1.xml.rels
      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
      nd_rels_ := Nyce_Xml.Make_Root_Node (doc_, 'Relationships', attrs_);

      nyce_xml.natr ('Id', 'rId1', attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition', attrs_);
      nyce_xml.attr ('Target', rep ('../pivotCache/pivotCacheDefinition:P1.xml', to_char(wb_.pivot_tables(pt_).cache_id)), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);

      Add1Xml (excel_, rep('xl/pivotTables/_rels/pivotTable:P1.xml.rels',pt_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);

   END LOOP;
END Finish_Pivot_Tables;


PROCEDURE Finish_Drawings_Rels (
   excel_ IN OUT NOCOPY BLOB,
   s_     IN            PLS_INTEGER )
IS
   img_id_  PLS_INTEGER;
   doc_     dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   attrs_   nyce_xml.xml_attrs_arr;
   nd_rels_ dbms_XmlDom.DomNode;
BEGIN

   IF wb_.sheets(s_).drawings.drawings_list.count = 0 THEN
      goto skip_drawings_rels;
   END IF;

   -- xl/drawings/_rels/drawing:P1.xml.rels
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rels_ := Nyce_Xml.Make_Root_Node (doc_, 'Relationships', attrs_);

   FOR dr_ IN 1 .. wb_.sheets(s_).drawings.drawings_list.count LOOP
      img_id_ := wb_.sheets(s_).drawings.drawings_list(dr_).img_id;
      nyce_xml.natr ('Id', 'rId' || to_char(dr_), attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', attrs_);
      nyce_xml.attr ('Target', rep ('../media/image:P1.:P2', to_char(img_id_), wb_.images(img_id_).extension), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
   END LOOP;

   Add1Xml (excel_, rep('xl/drawings/_rels/drawing:P1.xml.rels',s_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);

   <<skip_drawings_rels>>
   null;

END Finish_Drawings_Rels;

PROCEDURE Finish_Tables (
   excel_ IN OUT NOCOPY BLOB )
IS
   doc_     dbms_XmlDom.DomDocument;
   attrs_   nyce_xml.xml_attrs_arr;
   tbl_     tp_cell_range;
   nd_tbl_  dbms_XmlDom.DomNode;
   nd_tcls_ dbms_XmlDom.DomNode;
BEGIN

   IF wb_.tables_list.count = 0 THEN
      goto skip_tables;
   END IF;

   -- xl/tables/table:P1.xml
   FOR t_ IN 1 .. wb_.tables_list.count LOOP

      tbl_ := wb_.defined_names(wb_.tables_list(t_));

      doc_ := Dbms_XmlDom.newDomDocument;
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
      nyce_xml.attr ('id', to_char(t_), attrs_);
      nyce_xml.attr ('name', tbl_.defined_name, attrs_);
      nyce_xml.attr ('displayName', tbl_.defined_name, attrs_);
      nyce_xml.attr ('ref', Alfan_Range(tbl_), attrs_);
      nyce_xml.attr ('totalsRowShown', '0', attrs_);
      nd_tbl_ := Nyce_Xml.Make_Root_Node (doc_, 'table', attrs_);

      nyce_xml.natr ('ref', Alfan_Range(tbl_), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_tbl_, 'autoFilter', attrs_);

      nyce_xml.natr ('count', to_char(Range_Width(tbl_)), attrs_);
      nd_tcls_ := Nyce_Xml.Xml_Node (doc_, nd_tbl_, 'tableColumns', attrs_);
      FOR c_ IN 1 .. Range_Width(tbl_) LOOP
         nyce_xml.natr ('id', to_char(c_), attrs_);
         nyce_xml.attr ('name', tbl_.col_names(c_), attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_tcls_, 'tableColumn', attrs_);
      END LOOP;

      nyce_xml.natr ('name', tbl_.style, attrs_);
      nyce_xml.attr ('showFirstColumn', '0', attrs_);
      nyce_xml.attr ('showLastColumn', '0', attrs_);
      nyce_xml.attr ('showRowStripes', '1', attrs_);
      nyce_xml.attr ('showColumnStripes', '0', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_tbl_, 'tableStyleInfo', attrs_);

      Add1Xml (excel_, rep('xl/tables/table:P1.xml',t_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
      Dbms_XmlDom.freeDocument (doc_);

   END LOOP;

   <<skip_tables>>
   null;

END Finish_Tables;

PROCEDURE Finish_Worksheet (
   excel_ IN OUT NOCOPY BLOB,
   s_     IN            PLS_INTEGER )
IS
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   attrs_    nyce_xml.xml_attrs_arr;
   nd_ws_    dbms_XmlDom.DomNode;
   nd_svs_   dbms_XmlDom.DomNode;
   nd_sv_    dbms_XmlDom.DomNode;
   nd_cls_   dbms_XmlDom.DomNode;
   nd_sd_    dbms_XmlDom.DomNode;
   nd_r_     dbms_XmlDom.DomNode;
   nd_c_     dbms_XmlDom.DomNode;
   nd_mc_    dbms_XmlDom.DomNode;
   nd_dvs_   dbms_XmlDom.DomNode;
   nd_dv_    dbms_XmlDom.DomNode;
   nd_h_     dbms_XmlDom.DomNode;
   nd_tps_   dbms_XmlDom.DomNode;
   row_      PLS_INTEGER := wb_.sheets(s_).rows.first;
   col_      PLS_INTEGER;
   table_id_ PLS_INTEGER;
   col_min_  PLS_INTEGER := 16384;
   col_max_  PLS_INTEGER := 1;
   rel_      PLS_INTEGER := 1;
BEGIN

   WHILE row_ IS NOT null LOOP
      col_min_ := least (col_min_, wb_.sheets(s_).rows(row_).first);
      col_max_ := greatest (col_max_, wb_.sheets(s_).rows(row_).last);
      row_  := wb_.sheets(s_).rows.next(row_);
   END LOOP;

   -- xl/worksheets/sheet:P1.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   nyce_xml.attr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
   nyce_xml.attr ('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006', attrs_);
   nyce_xml.attr ('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', attrs_);
   nyce_xml.attr ('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision', attrs_);
   --nyce_xml.attr ('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', attrs_);
   nyce_xml.attr ('mc:Ignorable', 'x14ac', attrs_);
   nyce_xml.attr ('xr:uid', Get_Guid, attrs_);
   nd_ws_ := Nyce_Xml.Make_Root_Node (doc_, 'worksheet', attrs_);
   IF wb_.sheets(s_).tabcolor IS NOT null THEN
      nyce_xml.natr ('rgb', wb_.sheets(s_).tabcolor, attrs_);
      Nyce_Xml.Xml_Node (doc_, Nyce_Xml.Xml_Node(doc_,nd_ws_,'sheetPr'), 'tabColor', attrs_);
   END IF;

   nyce_xml.natr (
      'ref', Alfan_Range (
         col_tl_ => col_min_, row_tl_ => wb_.sheets(s_).rows.first,
         col_br_ => col_max_, row_br_ => wb_.sheets(s_).rows.last
      ), attrs_
   );
   Nyce_Xml.Xml_Node (doc_, nd_ws_, 'dimension', attrs_);

   nd_svs_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'sheetViews');
   nyce_xml.catr (attrs_);
   IF wb_.sheets(s_).grid_colour_ix IS NOT null THEN
      nyce_xml.attr ('defaultGridColor', '0', attrs_);
      nyce_xml.attr ('colorId', to_char(wb_.sheets(s_).grid_colour_ix), attrs_);
   END IF;
   nyce_xml.attr ('showGridLines',     '0', attrs_, not wb_.sheets(s_).show_gridlines);
   nyce_xml.attr ('showRowColHeaders', '0', attrs_, not wb_.sheets(s_).show_headers);
   nyce_xml.attr ('tabSelected',       '1', attrs_, s_=1);
   nyce_xml.attr ('workbookViewId',    '0', attrs_);
   nd_sv_  := Nyce_Xml.Xml_Node (doc_, nd_svs_, 'sheetView', attrs_);

   IF wb_.sheets(s_).freeze_rows + wb_.sheets(s_).freeze_cols > 0 THEN
      nyce_xml.natr ('activePane', 'bottomLeft', attrs_);
      nyce_xml.attr ('state', 'frozen', attrs_);
      IF wb_.sheets(s_).freeze_rows > 0 AND wb_.sheets(s_).freeze_cols > 0 THEN
         nyce_xml.attr ('xSplit', wb_.sheets(s_).freeze_cols, attrs_);
         nyce_xml.attr ('ySplit', wb_.sheets(s_).freeze_rows, attrs_);
         nyce_xml.attr ('topLeftCell', Alfan_Cell (wb_.sheets(s_).freeze_cols+1, wb_.sheets(s_).freeze_rows+1), attrs_);
      ELSIF wb_.sheets(s_).freeze_rows > 0 THEN
         nyce_xml.attr ('ySplit', wb_.sheets(s_).freeze_rows, attrs_);
         nyce_xml.attr ('topLeftCell', Alfan_Cell (1, wb_.sheets(s_).freeze_rows+1), attrs_);
      ELSIF wb_.sheets(s_).freeze_cols > 0 THEN
         nyce_xml.attr ('xSplit', wb_.sheets(s_).freeze_cols, attrs_);
         nyce_xml.attr ('topLeftCell', Alfan_Cell (wb_.sheets(s_).freeze_cols+1, 1), attrs_);
      END IF;
      Nyce_Xml.Xml_Node (doc_, nd_sv_, 'pane', attrs_);
   ELSE
      nyce_xml.natr ('activeCell', 'A1', attrs_);
      nyce_xml.attr ('sqref', 'A1', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_sv_, 'selection', attrs_);
   END IF;

   nyce_xml.natr ('defaultRowHeight', '15', attrs_);
   nyce_xml.attr ('x14ac:dyDescent', '0.25', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_ws_, 'sheetFormatPr', attrs_);

   IF wb_.sheets(s_).widths.count > 0 THEN
      nd_cls_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'cols');
      nyce_xml.catr (attrs_);
      col_ := wb_.sheets(s_).widths.first;
      WHILE col_ IS NOT null LOOP
         nyce_xml.natr ('min', col_, attrs_);
         nyce_xml.attr ('max', col_, attrs_);
         nyce_xml.attr ('width', to_char (wb_.sheets(s_).widths(col_), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'), attrs_);
         nyce_xml.attr ('customWidth', '1', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_cls_, 'col', attrs_);
         col_ := wb_.sheets(s_).widths.next(col_);
      END LOOP;
   END IF;

   -- <sheetData> goes here, included our calculated pivot tables
   nd_sd_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'sheetData');
   row_   := wb_.sheets(s_).rows.first;
   WHILE row_ IS NOT null LOOP
      nyce_xml.natr ('r', to_char(row_), attrs_);
      nyce_xml.attr ('spans', to_char(col_min_) || ':' || to_char(col_max_), attrs_);
      IF wb_.sheets(s_).row_fmts.exists(row_) AND wb_.sheets(s_).row_fmts(row_).height IS NOT null THEN
         nyce_xml.attr ('customHeight', '1', attrs_);
         nyce_xml.attr ('ht', to_char (wb_.sheets(s_).row_fmts(row_).height, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'), attrs_);
      END IF;
      nd_r_ := Nyce_Xml.Xml_Node (doc_, nd_sd_, 'row', attrs_);

      col_ := wb_.sheets(s_).rows(row_).first;
      WHILE col_ IS NOT null LOOP
         nyce_xml.natr ('r', Alfan_Cell (col_, row_), attrs_);
         IF wb_.sheets(s_).rows(row_)(col_).datatype IN (CELL_DT_STRING_, CELL_DT_HYPERLINK_) THEN
            nyce_xml.attr ('t', 's', attrs_);
         END IF;
         IF wb_.sheets(s_).rows(row_)(col_).style IS NOT null THEN
            nyce_xml.attr ('s', to_char(wb_.sheets(s_).rows(row_)(col_).style), attrs_);
         END IF;
         nd_c_ := Nyce_Xml.Xml_Node (doc_, nd_r_, 'c', attrs_);
         IF wb_.sheets(s_).rows(row_)(col_).formula_idx IS NOT null THEN
            Nyce_Xml.Xml_Text_Node (doc_, nd_c_, 'f', wb_.formulas(wb_.sheets(s_).rows(row_)(col_).formula_idx));
         END IF;
         Nyce_Xml.Xml_Text_Node (doc_, nd_c_, 'v', to_char(wb_.sheets(s_).rows(row_)(col_).value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,'));
         col_ := wb_.sheets(s_).rows(row_).next(col_);
      END LOOP;
      row_ := wb_.sheets(s_).rows.next(row_);
   END LOOP;

   FOR af_ IN 1 .. wb_.sheets(s_).autofilters.count LOOP
      nyce_xml.natr (
         'ref', Alfan_Range (
            col_tl_ => nvl (wb_.sheets(s_).autofilters(af_).column_start, col_min_),
            row_tl_ => nvl (wb_.sheets(s_).autofilters(af_).row_start, wb_.sheets(s_).rows.first),
            col_br_ => coalesce (wb_.sheets(s_).autofilters(af_).column_end, wb_.sheets(s_).autofilters(af_).column_start, col_max_),
            row_br_ => nvl (wb_.sheets(s_).autofilters(af_).row_end, wb_.sheets(s_).rows.last)
         ), attrs_
      );
      Nyce_Xml.Xml_Node (doc_, nd_ws_, 'autoFilter', attrs_);
   END LOOP;

   IF wb_.sheets(s_).mergecells.count > 0 THEN
      nyce_xml.natr ('count', to_char(wb_.sheets(s_).mergecells.count), attrs_);
      nd_mc_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'mergeCells', attrs_);
      FOR mg_ IN 1 .. wb_.sheets(s_).mergecells.count LOOP
         nyce_xml.natr ('ref', wb_.sheets(s_).mergecells(mg_), attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_mc_, 'mergeCell', attrs_);
      END LOOP;
   END IF;

   IF wb_.sheets(s_).validations.count > 0 THEN
      nyce_xml.natr ('count', to_char(wb_.sheets(s_).validations.count), attrs_);
      nd_dvs_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'dataValidations');

      FOR v_ IN wb_.sheets(s_).validations.count LOOP
         nyce_xml.natr ('type', wb_.sheets(s_).validations(v_).type, attrs_);
         nyce_xml.attr ('errorStyle', wb_.sheets(s_).validations(v_).errorstyle, attrs_);
         nyce_xml.attr ('allowBlank', CASE WHEN nvl(wb_.sheets(s_).validations(v_).allowBlank, true) THEN '1' ELSE '0' END, attrs_);
         nyce_xml.attr ('sqref', wb_.sheets(s_).validations(v_).sqref, attrs_);
         IF wb_.sheets(s_).validations(v_).prompt IS NOT null THEN
            nyce_xml.attr ('showInputMessage', '1', attrs_);
            nyce_xml.attr ('prompt', wb_.sheets(s_).validations(v_).prompt, attrs_);
            IF wb_.sheets(s_).validations(v_).title IS NOT null THEN
               nyce_xml.attr ('promptTitle', wb_.sheets(s_).validations(v_).title, attrs_);
            END IF;
         END IF;
         IF wb_.sheets(s_).validations(v_).showerrormessage THEN
            nyce_xml.attr ('showErrorMessage', '1', attrs_);
            IF wb_.sheets(s_).validations(v_).error_title IS NOT null THEN
               nyce_xml.attr ('errorTitle', wb_.sheets(s_).validations(v_).error_title, attrs_);
            END IF;
            IF wb_.sheets(s_).validations(v_).error_txt IS NOT null THEN
               nyce_xml.attr ('error', wb_.sheets(s_).validations(v_).error_txt, attrs_);
            END IF;
         END IF;
         nd_dv_ := Nyce_Xml.Xml_Node (doc_, nd_dvs_, 'dataValidation', attrs_);

         IF wb_.sheets(s_).validations(v_).formula1 IS NOT null THEN
            Nyce_Xml.Xml_Text_Node (doc_, nd_dv_, 'formula1', wb_.sheets(s_).validations(v_).formula1);
         END IF;
         IF wb_.sheets(s_).validations(v_).formula2 IS NOT null THEN
            Nyce_Xml.Xml_Text_Node (doc_, nd_dv_, 'formula2', wb_.sheets(s_).validations(v_).formula2);
         END IF;
      END LOOP;
   END IF;

   IF wb_.sheets(s_).hyperlinks.count > 0 THEN
      nd_h_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'hyperlinks');
      FOR h_ IN 1 .. wb_.sheets(s_).hyperlinks.count LOOP
         nyce_xml.natr ('ref', wb_.sheets(s_).hyperlinks(h_).cell, attrs_);
         nyce_xml.attr ('r:id', rep ('rId:P1', rel_), attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_h_, 'hyperlink', attrs_);
         wb_.sheets(s_).hyperlinks(h_).ws_rel := rel_;
         rel_ := rel_ + 1;
      END LOOP;
   END IF;

   nyce_xml.natr ('left', '0.7', attrs_);
   nyce_xml.attr ('right', '0.7', attrs_);
   nyce_xml.attr ('top', '0.75', attrs_);
   nyce_xml.attr ('bottom', '0.75', attrs_);
   nyce_xml.attr ('header', '0.3', attrs_);
   nyce_xml.attr ('footer', '0.3', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_ws_, 'pageMargins', attrs_);

   IF wb_.sheets(s_).drawings.drawings_list.count > 0 THEN
      nyce_xml.natr ('r:id', rep ('rId:P1', rel_), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ws_, 'drawing', attrs_);
      wb_.sheets(s_).drawings.ws_rel := rel_;
      rel_ := rel_ + 1;
   END IF;

   IF wb_.sheets(s_).comments.comments_list.count > 0 THEN
      nyce_xml.natr ('r:id', 'rId' || rel_, attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ws_, 'legacyDrawing', attrs_);
      wb_.sheets(s_).comments.ws_rel := rel_;
      rel_ := rel_ + 2; -- ../drawings/vmlDrawing1.vml + ../comments1.xml will be added to the rel sheet
   END IF;

   IF wb_.sheets(s_).tables_list.count > 0 THEN
      nyce_xml.natr ('count', to_char(wb_.sheets(s_).tables_list.count), attrs_);
      nd_tps_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'tableParts', attrs_);
      table_id_ := wb_.sheets(s_).tables_list.first;
      WHILE table_id_ IS NOT null LOOP
         nyce_xml.natr ('r:id', rep ('rId:P1', rel_), attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_tps_, 'tablePart', attrs_);
         wb_.defined_names(wb_.tables_list(table_id_)).ws_rel := rel_;
         rel_ := rel_ + 1;
         table_id_ := wb_.sheets(s_).tables_list.next(table_id_);
      END LOOP;
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
   nr_comments_   PLS_INTEGER := wb_.sheets(s_).comments.comments_list.count;
   nr_pivots_     PLS_INTEGER := wb_.sheets(s_).pivots_list.count;
   nr_drawings_   PLS_INTEGER := wb_.sheets(s_).drawings.drawings_list.count;
   nr_tables_     PLS_INTEGER := wb_.sheets(s_).tables_list.count;
   pivot_id_      PLS_INTEGER;
   table_id_      PLS_INTEGER;
   doc_           dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   attrs_         nyce_xml.xml_attrs_arr;
   nd_rels_       dbms_XmlDom.DomNode;
BEGIN

   IF nr_hyperlinks_ = 0 AND nr_comments_ = 0 AND nr_pivots_ = 0 AND nr_drawings_ = 0 AND nr_tables_ = 0 THEN
      goto skip_relationships;
   END IF;

   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships', attrs_);
   nd_rels_ := Nyce_Xml.Make_Root_Node (doc_, 'Relationships', attrs_);

   FOR h_ IN 1 .. nr_hyperlinks_ LOOP
      IF wb_.sheets(s_).hyperlinks(h_).url IS NOT null THEN
         nyce_xml.natr ('Id', rep ('rId:P1', wb_.sheets(s_).hyperlinks(h_).ws_rel), attrs_);
         nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', attrs_);
         nyce_xml.attr ('Target',  wb_.sheets(s_).hyperlinks(h_).url, attrs_);
         nyce_xml.attr ('TargetMode', 'External', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
         id_ := greatest (id_, wb_.sheets(s_).hyperlinks(h_).ws_rel);
      END IF;
   END LOOP;

   table_id_ := wb_.sheets(s_).tables_list.first;
   WHILE table_id_ IS NOT null LOOP
      nyce_xml.natr ('Id', 'rId' || to_char(wb_.defined_names(wb_.tables_list(table_id_)).ws_rel), attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', attrs_);
      nyce_xml.attr ('Target', rep('../tables/table:P1.xml', table_id_), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      id_ := greatest (id_, wb_.defined_names(wb_.tables_list(table_id_)).ws_rel);
      table_id_ := wb_.sheets(s_).tables_list.next(table_id_);
   END LOOP;

   IF nr_drawings_ > 0 THEN
      nyce_xml.natr ('Id', rep ('rId:P1', wb_.sheets(s_).drawings.ws_rel), attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', attrs_);
      nyce_xml.attr ('Target', rep ('../drawings/drawing:P1.xml', to_char(s_)), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      id_ := greatest (id_, wb_.sheets(s_).drawings.ws_rel);
   END IF;
   IF nr_comments_ > 0 THEN
      nyce_xml.natr ('Id', rep ('rId:P1', wb_.sheets(s_).comments.ws_rel), attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing', attrs_);
      nyce_xml.attr ('Target', rep ('../drawings/vmlDrawing:P1.vml', to_char(s_)), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);

      nyce_xml.natr ('Id', rep('rId:P1', wb_.sheets(s_).comments.ws_rel+1), attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', attrs_);
      nyce_xml.attr ('Target', rep ('../comments:P1.xml', to_char(s_)), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
      id_ := greatest (id_, wb_.sheets(s_).comments.ws_rel+1);
   END IF;

   -- The rId value of pivot tables does not have a corresponding rId value on
   -- the worksheet, so we'll just pick the next sequential number.
   FOR spid_ IN 1 .. wb_.sheets(s_).pivots_list.count LOOP
      id_ := id_ + 1;
      pivot_id_ := wb_.sheets(s_).pivots_list(spid_);
      nyce_xml.natr ('Id', rep('rId:P1', to_char(id_)), attrs_);
      nyce_xml.attr ('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable', attrs_);
      nyce_xml.attr ('Target', rep ('../pivotTables/pivotTable:P1.xml', pivot_id_), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_rels_, 'Relationship', attrs_);
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
   doc_      dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
   nd_ws_    dbms_XmlDom.DomNode;
   nd_tc_    dbms_XmlDom.DomNode;
   nd_fr_    dbms_XmlDom.DomNode;
   nd_to_    dbms_XmlDom.DomNode;
   nd_pi_    dbms_XmlDom.DomNode;
   nd_nv_    dbms_XmlDom.DomNode;
   nd_cn_    dbms_XmlDom.DomNode;
   nd_bf_    dbms_XmlDom.DomNode;
   nd_bl_    dbms_XmlDom.DomNode;
   nd_el_    dbms_XmlDom.DomNode;
   nd_et_    dbms_XmlDom.DomNode;
   attrs_    nyce_xml.xml_attrs_arr;
   drawing_  tp_drawing;
   to_col_   PLS_INTEGER;
   to_row_   PLS_INTEGER;
   col_ovfl_ NUMBER;
   row_ovfl_ NUMBER;
BEGIN

   IF wb_.sheets(s_).drawings.drawings_list.count = 0 THEN
      goto skip_drawings;
   END IF;

   -- xl/drawings/drawing:P1.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');
   nyce_xml.natr ('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', attrs_);
   nyce_xml.attr ('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main', attrs_);
   nd_ws_ := Nyce_Xml.Make_Root_Node (doc_, 'wsDr', 'xdr', attrs_);

   FOR dr_ IN 1 .. wb_.sheets(s_).drawings.drawings_list.count LOOP

      drawing_ := wb_.sheets(s_).drawings.drawings_list(dr_);
      Calc_Image_Col_And_Row (to_col_, to_row_, col_ovfl_, row_ovfl_, drawing_, s_);

      nyce_xml.natr ('editAs', 'oneCell', attrs_);
      nd_tc_ := Nyce_Xml.Xml_Node (doc_, nd_ws_, 'twoCellAnchor', 'xdr', attrs_);

      nd_fr_ := Nyce_Xml.Xml_Node (doc_, nd_tc_, 'from', 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_fr_, 'col', to_char(drawing_.col-1), 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_fr_, 'colOff', '0', 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_fr_, 'row', to_char(drawing_.row-1), 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_fr_, 'rowOff', '0', 'xdr');

      nd_to_ := Nyce_Xml.Xml_Node (doc_, nd_tc_, 'to', 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_to_, 'col', to_char(to_col_), 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_to_, 'colOff', to_char(col_ovfl_), 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_to_, 'row', to_char(to_row_), 'xdr');
      Nyce_Xml.Xml_Text_Node (doc_, nd_to_, 'rowOff', to_char(row_ovfl_), 'xdr');

      nd_pi_ := Nyce_Xml.Xml_Node (doc_, nd_tc_, 'pic', 'xdr');
      nd_nv_ := Nyce_Xml.Xml_Node (doc_, nd_pi_, 'nvPicPr', 'xdr');

      nyce_xml.natr ('id', '3', attrs_);
      nyce_xml.attr ('name', coalesce (drawing_.name, 'Picture '||dr_), attrs_);
      IF drawing_.title       IS NOT null THEN nyce_xml.attr('title', drawing_.title, attrs_); END IF;
      IF drawing_.description IS NOT null THEN nyce_xml.attr('descr', drawing_.description, attrs_); END IF;
      Nyce_Xml.Xml_Node (doc_, nd_nv_, 'cNvPr', 'xdr', attrs_);
      nd_cn_ := Nyce_Xml.Xml_Node (doc_, nd_nv_, 'cNvPicPr', 'xdr');

      nyce_xml.natr ('noChangeAspect', '1', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_cn_, 'picLocks', 'a', attrs_);

      nd_bf_ := Nyce_Xml.Xml_Node (doc_, nd_pi_, 'blipFill', 'xdr');

      nyce_xml.natr ('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrs_);
      nyce_xml.attr ('r:embed', rep ('rId:P1', to_char(dr_)), attrs_);
      nd_bl_ := Nyce_Xml.Xml_Node (doc_, nd_bf_, 'blip', 'a', attrs_);
      nd_et_ := Nyce_Xml.Xml_Node (doc_, nd_bl_, 'extLst', 'a');

      nyce_xml.natr ('uri', Get_Guid, attrs_);
      nd_el_ := Nyce_Xml.Xml_Node (doc_, nd_et_, 'ext', 'a', attrs_);

      nyce_xml.natr ('xmlns:a14', 'http://schemas.microsoft.com/office/drawing/2010/main', attrs_);
      nyce_xml.attr ('val', '0', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_el_, 'useLocalDpi', 'a14', attrs_);
      Nyce_Xml.Xml_Node (doc_, Nyce_Xml.Xml_Node(doc_,nd_bf_,'stretch','a'), 'fillRect', 'a');

      nyce_xml.natr ('prst', 'rect', attrs_);
      Nyce_Xml.Xml_Node (doc_, Nyce_Xml.Xml_Node(doc_,nd_pi_,'spPr','xdr'), 'prstGeom', 'a', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_tc_, 'clientData', 'xdr');

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
   attrs_         nyce_xml.xml_attrs_arr;
   nl_            VARCHAR2(2);
   comment_w_rem_ NUMBER;
   comment_h_     NUMBER;
   col_w_         NUMBER;
   colspan_       NUMBER;
BEGIN

   IF wb_.sheets(s_).comments.comments_list.count = 0 THEN
      goto skip_comments;
   END IF;

   FOR c_ IN 1 .. wb_.sheets(s_).comments.comments_list.count LOOP
      ws_authors_(wb_.sheets(s_).comments.comments_list(c_).author) := 0;
   END LOOP;

   -- xl/comments:P1.xml
   Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

   nyce_xml.natr ('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', attrs_);
   nd_cms_ := Nyce_Xml.Make_Root_Node (doc_, 'comments', attrs_);
   nd_aus_ := Nyce_Xml.Xml_Node (doc_, nd_cms_, 'authors');
   author_ := ws_authors_.first;
   WHILE author_ IS NOT null OR ws_authors_.next(author_) IS NOT null LOOP
      ws_authors_(author_) := au_count_;
      Nyce_Xml.Xml_Text_Node (doc_, nd_aus_, 'author', author_);
      au_count_  := au_count_ + 1;
      author_ := ws_authors_.next(author_);
   END LOOP;

   nd_cml_ := Nyce_Xml.Xml_Node (doc_, nd_cms_, 'commentList');
   FOR cm_ IN 1 .. wb_.sheets(s_).comments.comments_list.count LOOP
      nyce_xml.natr ('ref', Alfan_Cell (wb_.sheets(s_).comments.comments_list(cm_).column, wb_.sheets(s_).comments.comments_list(cm_).row), attrs_);
      nyce_xml.attr ('authorId', ws_authors_(wb_.sheets(s_).comments.comments_list(cm_).author), attrs_);
      nd_cm_ := Nyce_Xml.Xml_Node (doc_, nd_cml_, 'comment', attrs_);
      nd_tx_ := Nyce_Xml.Xml_Node (doc_, nd_cm_, 'text');
      IF wb_.sheets(s_).comments.comments_list(cm_).author IS NOT null THEN
         nd_r_  := Nyce_Xml.Xml_Node (doc_, nd_tx_, 'r');
         nd_pr_ := Nyce_Xml.Xml_Node (doc_, nd_r_, 'rPr');
         Nyce_Xml.Xml_Node (doc_, nd_pr_, 'b');

         nyce_xml.natr ('val', '9', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_pr_, 'sz', attrs_);

         nyce_xml.natr ('indexed', '81', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_pr_, 'color', attrs_);

         nyce_xml.natr ('val', 'Tahoma', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_pr_, 'rFont', attrs_);

         nyce_xml.natr ('val', '1', attrs_);
         Nyce_Xml.Xml_Node (doc_, nd_pr_, 'charset', attrs_);

         nyce_xml.natr ('xml:space', 'preserve', attrs_);
         Nyce_Xml.Xml_Text_Node (doc_, nd_r_, 't', wb_.sheets(s_).comments.comments_list(cm_).author, attrs_);
      END IF;
      nd_r_  := Nyce_Xml.Xml_Node (doc_, nd_tx_, 'r');
      nd_pr_ := Nyce_Xml.Xml_Node (doc_, nd_r_, 'rPr');

      nyce_xml.natr ('val', '9', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_pr_, 'sz', attrs_);

      nyce_xml.natr ('indexed', '81', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_pr_, 'color', attrs_);

      nyce_xml.natr ('val', 'Tahoma', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_pr_, 'rFont', attrs_);

      nyce_xml.natr ('val', '1', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_pr_, 'charset', attrs_);

      nyce_xml.natr ('xml:space', 'preserve', attrs_);
      nl_ := CASE WHEN wb_.sheets(s_).comments.comments_list(cm_).author IS NOT null THEN chr(13) || chr(10) END;
      Nyce_Xml.Xml_Text_Node (doc_, nd_r_, 't', nl_ || wb_.sheets(s_).comments.comments_list(cm_).text, attrs_);
   END LOOP;

   Add1Xml (excel_, rep('xl/comments:P1.xml',s_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);


   -- xl/drawings/vmlDrawing:P1.vml
   doc_ := Dbms_XmlDom.newDomDocument;

   nyce_xml.natr ('xmlns:v', 'urn:schemas-microsoft-com:vml', attrs_);
   nyce_xml.attr ('xmlns:o', 'urn:schemas-microsoft-com:office:office', attrs_);
   nyce_xml.attr ('xmlns:x', 'urn:schemas-microsoft-com:office:excel', attrs_);
   nd_xml_ := Nyce_Xml.Make_Root_Node (doc_, 'xml', attrs_);

   nyce_xml.natr ('v:ext', 'edit', attrs_);
   nd_sl_ := Nyce_Xml.Xml_Node (doc_, nd_xml_, 'shapelayout', 'o', attrs_);

   nyce_xml.natr ('v:ext', 'edit', attrs_);
   nyce_xml.attr ('data', '2', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_sl_, 'idmap', 'o', attrs_);

   nyce_xml.natr ('id', '_x0000_t202', attrs_);
   nyce_xml.attr ('coordsize', '21600,21600', attrs_);
   nyce_xml.attr ('o:spt', '202', attrs_);
   nyce_xml.attr ('path', 'm,l,21600r21600,l21600,xe', attrs_);
   nd_st_ := Nyce_Xml.Xml_Node (doc_, nd_xml_, 'shapetype', 'v', attrs_);

   nyce_xml.natr ('joinstyle', 'miter', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_st_, 'stroke', 'v', attrs_);

   nyce_xml.natr ('gradientshapeok', 't', attrs_);
   nyce_xml.attr ('o:connecttype', 'rect', attrs_);
   Nyce_Xml.Xml_Node (doc_, nd_st_, 'path', 'v', attrs_);

   FOR cm_ IN 1 .. wb_.sheets(s_).comments.comments_list.count LOOP

      nyce_xml.natr ('id', rep('_x0000_s:P1', to_char(cm_)), attrs_);
      nyce_xml.attr ('type', '#_x0000_t202', attrs_);
      nyce_xml.attr ('style', rep ('position:absolute;margin-left:35.25pt;margin-top:3pt;z-index::P1;visibility:hidden;', to_char(cm_)), attrs_);
      nyce_xml.attr ('fillcolor', '#ffffe1', attrs_);
      nyce_xml.attr ('o:insetmode', 'auto', attrs_);
      nd_sh_ := Nyce_Xml.Xml_Node (doc_, nd_xml_, 'shape', 'v', attrs_);

      nyce_xml.natr ('color2', '#ffffe1', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_sh_, 'fill', 'v', attrs_);

      nyce_xml.natr ('n', 't', attrs_);
      nyce_xml.attr ('color', 'black', attrs_);
      nyce_xml.attr ('obscured', 't', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_sh_, 'shadow', 'v', attrs_);

      nyce_xml.natr ('o:connecttype', 'none', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_sh_, 'path', 'v', attrs_);

      nyce_xml.natr ('style', 'mso-direction-alt:auto', attrs_);
      nd_tb_ := Nyce_Xml.Xml_Node (doc_, nd_sh_, 'textbox', 'v', attrs_);
      nyce_xml.attr ('style', 'text-align:left', attrs_);
      Nyce_Xml.Xml_Text_Node (doc_, nd_tb_, 'div', '', attrs_);

      nyce_xml.natr ('ObjectType', 'Note', attrs_);
      nd_cd_ := Nyce_Xml.Xml_Node (doc_, nd_sh_, 'ClientData', 'x', attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_cd_, 'MoveWithCells', 'x');
      Nyce_Xml.Xml_Node (doc_, nd_cd_, 'SizeWithCells', 'x');

      comment_w_rem_ := wb_.sheets(s_).comments.comments_list(cm_).width;
      comment_h_     := wb_.sheets(s_).comments.comments_list(cm_).height;
      colspan_       := 1;
      LOOP
         IF wb_.sheets(s_).widths.exists(wb_.sheets(s_).comments.comments_list(cm_).column+colspan_) THEN
            col_w_ := 256 * wb_.sheets(s_).widths(wb_.sheets(s_).comments.comments_list(cm_).column+colspan_);
            col_w_ := trunc((col_w_+18)/256*7); -- assume default 11 point Calibri
         ELSE
            col_w_ := 64;
         END IF;
         EXIT WHEN comment_w_rem_ < col_w_;
         colspan_       := colspan_ + 1;
         comment_w_rem_ := comment_w_rem_ - col_w_;
      END LOOP;
      Nyce_Xml.Xml_Text_Node (
         doc_, nd_cd_, 'Anchor',
         rep (
            ':P1,15,:P2,30,:P3,:P4,:P5,:P6',
            to_char(wb_.sheets(s_).comments.comments_list(cm_).column),
            to_char(wb_.sheets(s_).comments.comments_list(cm_).row),
            to_char(wb_.sheets(s_).comments.comments_list(cm_).column+colspan_-1),
            to_char(round(comment_w_rem_)),
            to_char(wb_.sheets(s_).comments.comments_list(cm_).row+1+trunc(comment_h_/20)),
            to_char(mod(comment_h_, 20))
         ), 'x'
      );
      Nyce_Xml.Xml_Text_Node (doc_, nd_cd_, 'AutoFill', 'False', 'x');
      Nyce_Xml.Xml_Text_Node (doc_, nd_cd_, 'Row', to_char(wb_.sheets(s_).comments.comments_list(cm_).row-1), 'x');
      Nyce_Xml.Xml_Text_Node (doc_, nd_cd_, 'Column', to_char(wb_.sheets(s_).comments.comments_list(cm_).column-1), 'x');
   END LOOP;

   Add1Xml (excel_, rep('xl/drawings/vmlDrawing:P1.vml',s_), Dbms_XmlDom.getXmlType(doc_).getClobVal);
   Dbms_XmlDom.freeDocument (doc_);


   << skip_comments >>
   null;

end Finish_Ws_Comments;


-----------
--- Encryption work goes here
--
--
$IF Nyce_Xlsx.DBMS_CRYPTO_INSTALLED_ $THEN

FUNCTION Encrypt_File (
   xl_file_ IN BLOB,
   user_pw_ IN VARCHAR2 ) RETURN BLOB
IS

   CLR_RED_   CONSTANT RAW(1) := hextoraw('00'); -- Red
   CLR_BLACK_ CONSTANT RAW(1) := hextoraw('01'); -- Black

   TYPE tp_children IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;
   TYPE tp_directory_entry IS RECORD (
      raw_name     RAW(64),
      entry_type   RAW(1),
      colour       RAW(1) := CLR_RED_,
      left         PLS_INTEGER := -1,
      right        PLS_INTEGER := -1,
      root         PLS_INTEGER := -1,
      children     tp_children,
      length       PLS_INTEGER := 0,
      first_sector PLS_INTEGER := 0 );
   TYPE tp_directory_list  IS TABLE OF tp_directory_entry INDEX BY PLS_INTEGER;
   TYPE tp_sector_ids IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;

   DIR_STORAGE_  CONSTANT RAW(1) := hexToRaw('01'); -- User storage
   DIR_STREAM_   CONSTANT RAW(1) := hexToRaw('02'); -- User stream
   DIR_ROOT_     CONSTANT RAW(1) := hexToRaw('05'); -- Root storage

   -- SAT = Sector Allocation Table
   FREE_SEC_ID_      CONSTANT PLS_INTEGER := -1; -- Free sector, may exist in the file, but is not part of any stream
   CHAIN_END_SEC_ID_ CONSTANT PLS_INTEGER := -2; -- Trailing SecID in a SecID chain
   SAT_SEC_ID_       CONSTANT PLS_INTEGER := -3; -- Sector is used by the sector allocation table

   primary_        RAW(200) := hexToRaw ('58000000010000004C0000007B00460046003900410033004600300033002D0035003600450046002D0034003600310033002D0042004400440035002D003500410034003100430031004400300037003200340036007D004E0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E0045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D00000001000000010000000100000000000000000000000000000004000000');
   se_data_space_  RAW(64)  := hexToRaw ('0800000001000000320000005300740072006F006E00670045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D000000');
   data_space_map_ RAW(112) := hexToRaw ('08000000010000006800000001000000000000002000000045006E0063007200790070007400650064005000610063006B00610067006500320000005300740072006F006E00670045006E006300720079007000740069006F006E004400610074006100530070006100630065000000');
   version_        RAW(76)  := hexToRaw ('3C0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E004400610074006100530070006100630065007300010000000100000001000000');

   encryption_info_   RAW(32767);
   encrypted_package_ BLOB;

   dir_list_     tp_directory_list;
   filesystem_   BLOB;
   short_stream_ BLOB;
   sctr_sz_      PLS_INTEGER := 512;
   ssctr_sz_     PLS_INTEGER := 64;
   ss_cutoff_    PLS_INTEGER := 4096;
   sc_id_        tp_sector_ids;
   ssc_id_       tp_sector_ids;
   msc_id_       tp_sector_ids;
   sector_count_ PLS_INTEGER;
   sector_diff_  PLS_INTEGER;
   root_dir_     PLS_INTEGER;
   storage_dir_  PLS_INTEGER;
   storage2_dir_ PLS_INTEGER;
   sorted_       BOOLEAN;
   dir_swap_     PLS_INTEGER;
   sectr_count_  PLS_INTEGER;
   sectrs_req_   PLS_INTEGER;
   header_       RAW(512);

   FUNCTION Is_Less (
      dir_entry1_ IN tp_directory_entry,
      dir_entry2_ IN tp_directory_entry ) RETURN BOOLEAN
   IS BEGIN
      RETURN CASE sign (Utl_Raw.Length(dir_entry1_.raw_name) - Utl_Raw.Length(dir_entry2_.raw_name))
         WHEN -1 THEN true
         WHEN  1 THEN false
         ELSE upper(utl_i18n.raw_to_char(dir_entry1_.raw_name, 'AL16UTF16LE')) -- what character set is this?
                 < upper(utl_i18n.raw_to_char(dir_entry2_.raw_name, 'AL16UTF16LE'))
      END;
   END Is_Less;

   FUNCTION Add_Dir_Entry (
      dir_name_   IN VARCHAR2,
      entry_type_ IN RAW,
      parent_     IN PLS_INTEGER := null,
      stream_     IN BLOB        := null,
      prefix_     IN RAW         := null ) RETURN PLS_INTEGER
   IS
      dir_count_ PLS_INTEGER := dir_list_.count;
      dir_entry_ tp_directory_entry;
   BEGIN
      dir_entry_.entry_type := entry_type_;
      dir_entry_.raw_name   := Utl_Raw.Concat(prefix_, Utl_I18n.String_To_Raw(dir_name_,'AL16UTF16LE'));
      IF parent_ IS NOT null THEN
         dir_list_(parent_).children(dir_list_(parent_).children.count) := dir_count_;
      END IF;
      IF entry_type_ = DIR_STREAM_ THEN
         dir_entry_.length := Dbms_Lob.getLength (stream_);
         IF dir_entry_.length >= ss_cutoff_ THEN
            Dbms_Lob.Append (filesystem_, stream_);
            IF mod (dir_entry_.length, sctr_sz_) > 0 THEN
               Dbms_Lob.writeAppend (
                  filesystem_,
                  sctr_sz_ - mod(dir_entry_.length, sctr_sz_),
                  Utl_Raw.Copies('00', sctr_sz_)
               );
            END IF;
            dir_entry_.first_sector := sc_id_.count;
            FOR i_ IN sc_id_.count .. sc_id_.count + trunc((dir_entry_.length-1)/sctr_sz_)-1 LOOP
               sc_id_(i_) := i_ + 1;
            END LOOP;
            sc_id_(sc_id_.count) := CHAIN_END_SEC_ID_;
         ELSE
            Dbms_Lob.Append (short_stream_, stream_);
            IF mod(dir_entry_.length, ssctr_sz_) > 0 THEN
               Dbms_Lob.writeAppend (short_stream_, ssctr_sz_ - mod(dir_entry_.length, ssctr_sz_), utl_raw.copies('00', ssctr_sz_));
            END IF;
            dir_entry_.first_sector := ssc_id_.count;
            FOR i_ IN ssc_id_.count .. ssc_id_.count + trunc((dir_entry_.length-1)/ssctr_sz_)-1 LOOP
               ssc_id_(i_) := i_ + 1;
            END LOOP;
            ssc_id_(ssc_id_.count) := CHAIN_END_SEC_ID_;
         END IF;
      END IF;
      dir_list_(dir_count_) := dir_entry_;
      RETURN dir_count_;
   END Add_Dir_Entry;

   PROCEDURE Add_Dir_Entry (
      dir_name_   IN VARCHAR2,
      entry_type_ IN RAW,
      parent_     IN PLS_INTEGER := null,
      stream_     IN BLOB        := null,
      prefix_     IN RAW         := null )
   IS
      throw_nr_ PLS_INTEGER;
   BEGIN
      throw_nr_ := Add_Dir_Entry (dir_name_, entry_type_, parent_, stream_, prefix_);
   END Add_Dir_Entry;

   PROCEDURE Do_Encryption (
      package_ IN OUT NOCOPY BLOB )
   IS
      -- bk = block-key
      ENCR_VER_HASH_INPUT_BK_ CONSTANT RAW(8) := hexToRaw ('fea7d2763b4b9e79'); -- encrVerifierHashInputBlockKey
      ENCR_VER_HASH_VALUE_BK_ CONSTANT RAW(8) := hexToRaw ('d7aa0f6d3061344e'); -- encrVerifierHashValueBlockKey
      ENCR_KEY_VAL_BK_        CONSTANT RAW(8) := hexToRaw ('146e0be7abacd0d6'); -- encryptedKeyValueBlockKey
      ENCR_INTEGRITY_SALT_BK_ CONSTANT RAW(8) := hexToRaw ('5fb2ad010cb9e1f6'); -- encrIntegritySaltBlockKey
      ENCR_INTEGRITY_HMAV_BK_ CONSTANT RAW(8) := hexToRaw ('a0677f02b22c8433'); -- encrIntegrityHmacValueBlocKkey

      ALGO_              CONSTANT PLS_INTEGER := Dbms_Crypto.ENCRYPT_AES + Dbms_Crypto.CHAIN_CBC + Dbms_Crypto.PAD_ZERO; -- c_algo
      KEY_BITS_          CONSTANT PLS_INTEGER := 256 / 8;

      hash_sh1_          CONSTANT PLS_INTEGER := Dbms_Crypto.Hash_Sh1;
      hmac_sh1_          CONSTANT PLS_INTEGER := Dbms_Crypto.Hmac_Sh1;
      HASH_ALGO_         CONSTANT VARCHAR2(4) := 'SHA1';
      HASH_LEN_          CONSTANT PLS_INTEGER := Utl_Raw.Length (Dbms_Crypto.Hash('00', hash_sh1_));
      BLOCK_SIZE_        CONSTANT PLS_INTEGER := 16;
      SPIN_COUNT_        CONSTANT PLS_INTEGER := 1000;
      SALT_SIZE_         CONSTANT PLS_INTEGER := 16;
      SALT_              CONSTANT RAW(3999)   := Dbms_Crypto.randomBytes (SALT_SIZE_);
      DATA_SALT_         CONSTANT RAW(3999)   := Dbms_Crypto.randomBytes (SALT_SIZE_);
      PW_                CONSTANT RAW(32767)  := Utl_i18n.String_To_Raw (user_pw_, 'AL16UTF16LE');
      xl_size_           CONSTANT INTEGER     := Dbms_Lob.getLength (xl_file_);

      decrypted_key_val_ RAW(100)    := Dbms_Crypto.randomBytes(KEY_BITS_);
      salt_raw_          RAW(100)    := Dbms_Crypto.randomBytes(HASH_LEN_);
      last_block_        PLS_INTEGER := trunc ((xl_size_-1)/4096);
      r_key_             RAW(100);
      r_inp_             RAW(100);
      iv_raw_            RAW(100);
      enc_key_val_       VARCHAR2(100);
      xl_block_          RAW(4096);
      mac_               RAW(100);
      enc_hmac_key_      VARCHAR2(100);
      enc_hmac_value_    VARCHAR2(100);
      enc_vrifr_input_   VARCHAR2(100);
      enc_vrifr_value_   VARCHAR2(100);
      hash_raw_          RAW(100);

      doc_    dbms_XmlDom.DomDocument := Dbms_XmlDom.newDomDocument;
      nd_enc_ dbms_XmlDom.DomNode;
      nd_ke_  dbms_XmlDom.DomNode;
      attrs_  nyce_xml.xml_attrs_arr;

      FUNCTION Generate_Key (
         block_key_ IN RAW ) RETURN RAW
      IS
         hash_buf_ RAW(1000);
      BEGIN
         hash_buf_ := Dbms_Crypto.Hash (Utl_Raw.Concat(SALT_,PW_), hash_sh1_);
         FOR i_ IN 0 .. SPIN_COUNT_ - 1 LOOP
            hash_buf_ := Dbms_Crypto.Hash (Utl_Raw.Concat(Little_Endian(i_),hash_buf_), hash_sh1_);
         END LOOP;
         hash_buf_ := Dbms_Crypto.Hash (Utl_Raw.Concat(hash_buf_,block_key_), hash_sh1_);
         IF HASH_LEN_ < KEY_BITS_ THEN
            hash_buf_ := Utl_Raw.Concat (hash_buf_, utl_raw.copies(hextoraw('36'), KEY_BITS_));
         END IF;
         RETURN Utl_Raw.Substr (hash_buf_, 1, KEY_BITS_);
      END Generate_Key;

   BEGIN
      r_key_       := Generate_Key (ENCR_KEY_VAL_BK_);
      iv_raw_      := Dbms_Crypto.Encrypt (decrypted_key_val_, ALGO_, r_key_, salt_);
      enc_key_val_ := Utl_Raw.Cast_To_Varchar2 (Utl_Encode.Base64_Encode(iv_raw_));
      package_     := Little_Endian (xl_size_, 8);
      FOR i_ IN 0 .. last_block_ LOOP
         iv_raw_ := Dbms_Crypto.Hash (Utl_Raw.Concat(DATA_SALT_,Little_Endian(i_)), hash_sh1_);
         IF HASH_LEN_ < BLOCK_SIZE_ THEN
            iv_raw_ := Utl_Raw.Concat (iv_raw_, Utl_Raw.Copies(hexToRaw('36'), BLOCK_SIZE_));
         END IF;
         iv_raw_   := Utl_Raw.Substr (iv_raw_, 1, BLOCK_SIZE_);
         xl_block_ := Dbms_Lob.Substr (xl_file_, 4096, i_*4096 + 1);
         IF i_ = last_block_ AND mod (Utl_Raw.Length(xl_block_), BLOCK_SIZE_) != 0 THEN
            xl_block_ := Utl_Raw.Concat (
               xl_block_, Utl_Raw.Copies (
                  'FF', BLOCK_SIZE_-mod(Utl_Raw.Length(xl_block_), BLOCK_SIZE_)
               )
            );
         END IF;
         Dbms_Lob.Append (
            package_, Dbms_Crypto.Encrypt (xl_block_, ALGO_, decrypted_key_val_, iv_raw_)
         );
      END LOOP;
      mac_    := Dbms_Crypto.Mac (package_, hmac_sh1_, salt_raw_);
      iv_raw_ := Dbms_Crypto.Hash (Utl_Raw.Concat (DATA_SALT_, ENCR_INTEGRITY_SALT_BK_), hash_sh1_);
      IF Utl_Raw.Length(iv_raw_) < BLOCK_SIZE_ THEN
         iv_raw_ := Utl_Raw.Concat (iv_raw_, Utl_Raw.Copies (hexToRaw('00'), BLOCK_SIZE_));
      END IF;
      iv_raw_   := Utl_Raw.Substr (iv_raw_, 1, BLOCK_SIZE_);
      salt_raw_ := Dbms_Crypto.Encrypt (salt_raw_, ALGO_, decrypted_key_val_, iv_raw_);
      enc_hmac_key_ := Utl_Raw.Cast_To_Varchar2 (Utl_Encode.Base64_Encode(salt_raw_));
      iv_raw_   := Dbms_Crypto.Hash (Utl_Raw.Concat (DATA_SALT_, ENCR_INTEGRITY_HMAV_BK_), hash_sh1_);
      IF Utl_Raw.Length(iv_raw_) < BLOCK_SIZE_ THEN
         iv_raw_ := Utl_Raw.Concat (iv_raw_, Utl_Raw.Copies(hexToRaw('00'), BLOCK_SIZE_));
      END IF;

      iv_raw_   := Utl_Raw.Substr (iv_raw_, 1, BLOCK_SIZE_);
      hash_raw_ := Dbms_Crypto.Encrypt (mac_, ALGO_, decrypted_key_val_, iv_raw_);
      enc_hmac_value_ := Utl_Raw.Cast_To_Varchar2 (Utl_Encode.Base64_Encode(hash_raw_));

      r_inp_ := Dbms_Crypto.randomBytes (SALT_SIZE_);
      r_key_ := Generate_Key (ENCR_VER_HASH_INPUT_BK_);
      hash_raw_ := Dbms_Crypto.Encrypt (r_inp_, ALGO_, r_key_, SALT_);
      enc_vrifr_input_ := Utl_Raw.Cast_To_Varchar2 (Utl_Encode.Base64_Encode(hash_raw_));
      r_key_ := Generate_Key (ENCR_VER_HASH_VALUE_BK_);
      r_inp_ := Dbms_Crypto.Hash (r_inp_, hash_sh1_);
      hash_raw_ := Dbms_Crypto.Encrypt (r_inp_, ALGO_, r_key_, SALT_);
      enc_vrifr_value_ := Utl_Raw.Cast_To_Varchar2 (Utl_Encode.Base64_Encode(hash_raw_));

      -- then generate the XML
      Dbms_XmlDom.setVersion (doc_, '1.0" encoding="UTF-8" standalone="yes');

      nyce_xml.natr ('xmlns', 'http://schemas.microsoft.com/office/2006/encryption', attrs_);
      nyce_xml.attr ('xmlns:p', 'http://schemas.microsoft.com/office/2006/keyEncryptor/password', attrs_);
      nd_enc_ := Nyce_Xml.Make_Root_Node (doc_, 'encryption', attrs_);

      nyce_xml.natr ('saltSize',        to_char(SALT_SIZE_),  attrs_);
      nyce_xml.attr ('blockSize',       to_char(BLOCK_SIZE_), attrs_);
      nyce_xml.attr ('keyBits',         to_char(KEY_BITS_*8), attrs_);
      nyce_xml.attr ('hashSize',        to_char(HASH_LEN_),   attrs_);
      nyce_xml.attr ('cipherAlgorithm', 'AES',                attrs_);
      nyce_xml.attr ('cipherChaining',  'ChainingModeCBC',    attrs_);
      nyce_xml.attr ('hashAlgorithm',   HASH_ALGO_,           attrs_);
      nyce_xml.attr ('saltValue', Utl_Raw.Cast_To_Varchar2(Utl_Encode.Base64_Encode(DATA_SALT_)), attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_enc_, 'keyData', attrs_);

      nyce_xml.natr ('encryptedHmacKey',   enc_hmac_key_,   attrs_);
      nyce_xml.attr ('encryptedHmacValue', enc_hmac_value_, attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_enc_, 'dataIntegrity', attrs_);

      nyce_xml.natr ('uri', 'http://schemas.microsoft.com/office/2006/keyEncryptor/password', attrs_);
      nd_ke_ := Nyce_Xml.Xml_Node (doc_, nd_enc_, 'keyEncryptors/keyEncryptor', attrs_);

      nyce_xml.natr ('spinCount',                  to_char(SPIN_COUNT_), attrs_);
      nyce_xml.attr ('saltSize',                   to_char(SALT_SIZE_),  attrs_);
      nyce_xml.attr ('blockSize',                  to_char(BLOCK_SIZE_), attrs_);
      nyce_xml.attr ('keyBits',                    to_char(KEY_BITS_*8), attrs_);
      nyce_xml.attr ('hashSize',                   to_char(HASH_LEN_),   attrs_);
      nyce_xml.attr ('cipherAlgorithm',            'AES',                attrs_);
      nyce_xml.attr ('cipherChaining',             'ChainingModeCBC',    attrs_);
      nyce_xml.attr ('hashAlgorithm',              HASH_ALGO_,           attrs_);
      nyce_xml.attr ('saltValue', Utl_Raw.Cast_To_Varchar2(Utl_Encode.Base64_Encode(SALT_)), attrs_);
      nyce_xml.attr ('encryptedVerifierHashInput', enc_vrifr_input_,     attrs_);
      nyce_xml.attr ('encryptedVerifierHashValue', enc_vrifr_value_,     attrs_);
      nyce_xml.attr ('encryptedKeyValue',          enc_key_val_,         attrs_);
      Nyce_Xml.Xml_Node (doc_, nd_ke_, 'p:encryptedKey', attrs_);
      encryption_info_ := Utl_Raw.Concat (
         hexToRaw('0400040040000000'), Utl_Raw.Cast_To_Raw (Dbms_XmlDom.getXmlType(doc_).getClobVal)
      );
   END Do_Encryption;

BEGIN

   Do_Encryption (encrypted_package_);

   filesystem_ := Utl_Raw.Copies ('00', sctr_sz_);
   Dbms_Lob.createTemporary (short_stream_, true);
   root_dir_ := Add_Dir_Entry ('Root Entry', DIR_ROOT_);
   Add_Dir_Entry ('EncryptedPackage', DIR_STREAM_, root_dir_, encrypted_package_);
   storage_dir_ := Add_Dir_Entry ('DataSpaces', DIR_STORAGE_, root_dir_, prefix_ => '0600');
   Add_Dir_Entry ('Version', DIR_STREAM_, storage_dir_, version_);
   Add_Dir_Entry ('DataSpaceMap', DIR_STREAM_, storage_dir_, data_space_map_);
   storage2_dir_ := Add_Dir_Entry ('DataSpaceInfo', DIR_STORAGE_, storage_dir_);
   Add_Dir_Entry ('StrongEncryptionDataSpace', DIR_STREAM_, storage2_dir_, se_data_space_);
   Add_Dir_Entry ('TransformInfo', DIR_STORAGE_, storage_dir_);
   storage2_dir_ := Add_Dir_Entry ('StrongEncryptionTransform', DIR_STORAGE_, storage2_dir_);
   Add_Dir_Entry ('Primary', DIR_STREAM_, storage2_dir_, primary_, prefix_ => '0600');
   Add_Dir_Entry ('EncryptionInfo', DIR_STREAM_, root_dir_, encryption_info_);
   Dbms_Lob.freeTemporary (encrypted_package_);

   -- write the short sector stream
   Dbms_Lob.Append (filesystem_, short_stream_);
   IF mod(Dbms_Lob.getLength(short_stream_), sctr_sz_) > 0 THEN
      Dbms_Lob.writeAppend (
         filesystem_, sctr_sz_ - mod(Dbms_Lob.getLength(short_stream_), sctr_sz_),
         Utl_Raw.Copies ('00', sctr_sz_)
      );
   END IF;
   dir_list_(0).length       := Dbms_Lob.getLength (short_stream_);
   dir_list_(0).first_sector := sc_id_.count;
   FOR i_ IN sc_id_.count .. sc_id_.count + trunc((Dbms_Lob.getLength(short_stream_)-1)/sctr_sz_)-1 LOOP
      sc_id_(i_) := i_ + 1;
   END LOOP;
   sc_id_(sc_id_.count) := CHAIN_END_SEC_ID_;
   --
   -- write the ssat
   FOR i_ IN 0 .. ssc_id_.count - 1 LOOP
      Dbms_Lob.writeAppend (filesystem_, 4, Little_Endian(ssc_id_(i_)));
   END LOOP;
   IF mod (ssc_id_.count*4, sctr_sz_) > 0 THEN
      Dbms_Lob.writeAppend (
         filesystem_, sctr_sz_ - mod(ssc_id_.count*4, sctr_sz_),
         Utl_Raw.Copies (Little_Endian(FREE_SEC_ID_), sctr_sz_)
      );
   END IF;
   sector_count_ := sc_id_.count;
   FOR i_ IN sc_id_.count .. sc_id_.count + trunc((ssc_id_.count*4-1)/sctr_sz_)-1 LOOP
      sc_id_(i_) := i_ + 1;
   END LOOP;
   sc_id_(sc_id_.count) := CHAIN_END_SEC_ID_;
   sector_diff_ := sc_id_.count - sector_count_;

   FOR i_ IN 0 .. dir_list_.last LOOP
      IF dir_list_(i_).children.count = 1 THEN
         dir_list_(i_).root := dir_list_(i_).children(0);
         dir_list_(dir_list_(i_).children(0)).colour := CLR_BLACK_;
      ELSIF dir_list_(i_).children.count > 1 THEN
         sorted_ := false;
         WHILE not sorted_ LOOP
            sorted_ := true;
            FOR j_ IN 0 .. dir_list_(i_).children.count - 2 LOOP
               IF Is_Less (dir_list_(dir_list_(i_).children(j_+1)), dir_list_(dir_list_(i_).children(j_))) THEN
                  dir_swap_               := dir_list_(i_).children(j_);
                  dir_list_(i_).children(j_)   := dir_list_(i_).children(j_+1);
                  dir_list_(i_).children(j_+1) := dir_swap_;
                  sorted_ := false;
               END IF;
            END LOOP;
         END LOOP;
         dir_swap_                   := dir_list_(i_).children(1);
         dir_list_(i_).root          := dir_swap_;
         dir_list_(dir_swap_).left   := dir_list_(i_).children(0);
         dir_list_(dir_swap_).colour := CLR_BLACK_;
         IF dir_list_(i_).children.count > 2 THEN
            dir_list_(dir_swap_).right := dir_list_(i_).children(2);
            IF dir_list_(i_).children.count > 3 THEN
               dir_list_(dir_list_(i_).children(2)).right  := dir_list_(i_).children(3);
               dir_list_(dir_list_(i_).children(0)).colour := CLR_BLACK_;
               dir_list_(dir_list_(i_).children(2)).colour := CLR_BLACK_;
            END IF;
         END IF;
      END IF;
   END LOOP;
   
   FOR i_ IN 0 .. dir_list_.count - 1 LOOP
      Dbms_Lob.writeAppend (
         filesystem_, 128,
         Utl_Raw.Concat (
            Utl_Raw.Overlay ('00', dir_list_(i_).raw_name, 64),
            Little_Endian (Utl_Raw.Length(dir_list_(i_).raw_name)+2, 2),
            dir_list_(i_).entry_type,
            dir_list_(i_).colour,
            Little_Endian(dir_list_(i_).left),
            Little_Endian(dir_list_(i_).right),
            Little_Endian(dir_list_(i_).root),
            Utl_Raw.Copies('00', 36),
            Little_Endian(dir_list_(i_).first_sector),
            Little_Endian(dir_list_(i_).length),
            Utl_Raw.Copies('00',4)
         )
      );
   END LOOP;
   IF mod (dir_list_.count*128, sctr_sz_) > 0 THEN
      Dbms_Lob.writeAppend (filesystem_, sctr_sz_-mod(dir_list_.count*128, sctr_sz_), Utl_Raw.Copies('00',sctr_sz_));
   END IF;
   sectr_count_ := sc_id_.count;
   FOR i_ IN sectr_count_ .. sectr_count_ + trunc((dir_list_.count*128-1)/sctr_sz_)-1 LOOP
      sc_id_(i_) := i_ + 1;
   END LOOP;
   sc_id_(sc_id_.count) := CHAIN_END_SEC_ID_;
   --
   -- write the sat
   sectrs_req_ := floor (sc_id_.count* 4/sctr_sz_);
   FOR i_ IN 0 .. sectrs_req_ LOOP
      msc_id_(msc_id_.count) := sc_id_.count;
      sc_id_(sc_id_.count)   := SAT_SEC_ID_;
   END LOOP;
   IF sectrs_req_ != floor (sc_id_.count* 4/sctr_sz_) THEN
      msc_id_(msc_id_.count) := sc_id_.count;
      sc_id_(sc_id_.count)   := SAT_SEC_ID_;
   END IF;
   FOR i_ IN 0 .. sc_id_.count - 1 LOOP
      Dbms_Lob.writeAppend (filesystem_, 4, Little_Endian(sc_id_(i_)));
   END LOOP;
   IF mod(sc_id_.count*4, sctr_sz_) > 0 THEN
      Dbms_Lob.writeAppend (
         filesystem_, sctr_sz_-mod(sc_id_.count*4, sctr_sz_),
         Utl_Raw.Copies(Little_Endian(FREE_SEC_ID_), sctr_sz_)
      );
   END IF;
   header_ := Utl_Raw.Concat (
      hexToRaw ('D0CF11E0A1B11AE1'), Utl_Raw.Copies ('00', 16),
      hexToRaw ('3E000300'), hexToRaw ('FEFF'),
      Little_Endian (round(log(2,sctr_sz_)), 2),
      Little_Endian (round(log(2,ssctr_sz_)), 2),
      Utl_Raw.Copies ('00', 10), Little_Endian (msc_id_.count),
      Little_Endian (sectr_count_), Utl_Raw.Copies ('00', 4),
      Little_Endian (ss_cutoff_), Little_Endian (sector_count_)
   );
   header_ := Utl_Raw.Concat (
      header_, Little_Endian(sector_diff_),
      Little_Endian(CHAIN_END_SEC_ID_), Utl_Raw.Copies('00',4)
   );
   FOR i_ IN 0 .. msc_id_.count - 1 LOOP
      header_ := Utl_Raw.Concat (header_, Little_Endian(msc_id_(i_)));
   END LOOP;
   header_ := Utl_Raw.Concat (header_, Utl_Raw.Copies (Little_Endian(FREE_SEC_ID_), 109-msc_id_.count));
   Dbms_Lob.Copy (filesystem_, header_, 512, 1, 1);
   Dbms_Lob.freeTemporary (short_stream_);
   RETURN filesystem_;

END Encrypt_File;
$END

FUNCTION Finish (
   pw_ IN VARCHAR2 := '' ) RETURN BLOB
IS
   excel_ BLOB;
   s_     PLS_INTEGER;
BEGIN

   -- Pad out the Pivot Cache before doing any Excel generation.  Pivot tables
   -- will need this data in a moment...
   Build_Pivot_Caches_And_Tables;

   Dbms_Lob.createTemporary (excel_, true);

   -- We need to sort out the Pivots first, because tables will inject data to
   -- sheets, which in turn will create some additional shared strings and the
   -- like.  All this needs to be modelled in PL/SQL data-structures before we
   -- go about building out the XML for the various "parts"
   Finish_Pivot_Caches (excel_);            -- xl/pivotCache/pivotCacheDefinition[1].xml
   Finish_Pivot_Tables (excel_);            -- xl/pivotTables/pivotTable[1].xml

   Finish_Content_Types (excel_);           -- [Content_Types].xml
   Finish_docProps (excel_);                -- docProps/core.xml
   Finish_Rels (excel_);                    -- _rels/.rels
   Finish_Shared_Strings (excel_);          -- xl/sharedStrings.xml
   Finish_Styles (excel_);                  -- xl/styles.xml
   Finish_Theme (excel_);                   -- xl/theme/theme1.xml
   Finish_Workbook (excel_);                -- xl/workbook.xml
   Finish_Workbook_Rels (excel_);           -- xl/_rels/workbook.xml.rels
   Finish_Media (excel_);                   -- xl/media/image:P1.[bmp/gif/jpg/png]
   Finish_Tables (excel_);                  -- xl/tables/table:P1.xml

   s_ := wb_.sheets.first;
   WHILE s_ IS not null LOOP
      Finish_Worksheet (excel_, s_);        -- xl/worksheets/sheet:P1.xml
      Finish_Ws_Relationships (excel_, s_); -- xl/worksheets/_rels/sheet:P1.xml.rels
      Finish_Ws_Drawings (excel_, s_);      -- xl/drawings/drawing:P1.xml
      Finish_Drawings_Rels (excel_, s_);    -- xl/drawings/_rels/drawing:P1.xml.rels
      Finish_Ws_Comments (excel_, s_);      -- xl/drawings/vmlDrawing:P1.vml
      s_ := wb_.sheets.next(s_);
   END LOOP;

   Finish_Zip (excel_);
   Clear_Workbook;

   $IF Nyce_Xlsx.DBMS_CRYPTO_INSTALLED_ $THEN
      IF pw_ IS NOT null THEN
         excel_ := Encrypt_File (excel_, pw_);
      END IF;
   $END
   RETURN excel_;

END Finish;

PROCEDURE Save (
   directory_ IN VARCHAR2,
   filename_  IN VARCHAR2,
   pw_        IN VARCHAR2 := '' )
IS BEGIN
   Blob2File (Finish(pw_), directory_, filename_);
END Save;

PROCEDURE Save (
   xl_blob_   IN BLOB,
   directory_ IN VARCHAR2,
   filename_  IN VARCHAR2 )
IS BEGIN
   Blob2File (xl_blob_, directory_, filename_);
END Save;

-----
-- Query2Sheet()
-- Query2SeehtAndAutofilter()
-- Query2Table()
--   This collection of functions is the quickest way of putting data onto the
--   Excel sheet.  col_fmts_ allows us to define one numFmt for each column of
--   the data-source.  You can leave the collection sparse if some columns are
--   not in need of formatting.  It'll default back to those you set earlier.
--
PROCEDURE Query2Sheet (
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   crs_         IN OUT NOCOPY INTEGER,
   col_headers_ IN BOOLEAN        := true,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   col_numFmts_   tp_numFmt_cols := col_fmts_;
   sh_            PLS_INTEGER    := CASE WHEN sheet_ IS null THEN New_Sheet ELSE sheet_ END;
   curr_row_      PLS_INTEGER    := nvl (row_pos_, 1);
   curr_col_      PLS_INTEGER    := nvl (col_pos_-1, 0); -- each offset is at least 1
   align_hz_      VARCHAR2(16);
   align_right_   BOOLEAN        := false;
   align_ctrcont_ BOOLEAN        := false;
   desc_tab_      dbms_sql.desc_tab2;
   d_tab_         dbms_sql.date_table;
   n_tab_         dbms_sql.number_table;
   v_tab_         dbms_sql.varchar2_table;
   data_len_      NUMBER;
   bulk_sz_       PLS_INTEGER := 200;
   rows_fetched_  INTEGER;
   widths_        tp_widths; -- coarsly, we auto-set the column widths to fit the data
   ix_            NUMBER;
BEGIN

   IF curr_row_ < 1 OR curr_col_ < 0 THEN
      Raise_App_Error ('Table must be placed on row >=1, and on column >=0');
   END IF;

   -- First set up the title cell, which lies across the top of the table, and
   -- merges cells across the table's width
   IF title_ IS NOT null THEN
      curr_row_ := curr_row_ + 1;
      IF title_xfId_ IS NOT null AND wb_.cellXfs.exists(title_xfId_) THEN
         align_hz_ := lower (wb_.cellXfs(title_xfId_).alignment.horizontal);
         align_right_   := align_hz_ = 'right';
         align_ctrcont_ := align_hz_ = 'centercontinuous';
      END IF;
   END IF;

   -- Then sort the data grid itself (with or without column headers)
   Dbms_Sql.Describe_Columns2 (crs_, col_count_, desc_tab_);

   FOR col_ IN 1 .. col_count_ LOOP
      IF title_ IS NOT null THEN
         IF (col_=1 AND not align_right_) OR (col_=col_count_ AND align_right_) THEN
            Cell (curr_col_+col_, curr_row_-1, value_ => title_, sheet_ => sh_);
            wb_.sheets(sh_).rows(curr_row_-1)(curr_col_+col_).style := title_xfId_;
         ELSIF align_ctrcont_ THEN
            CellB (curr_col_+col_, curr_row_-1, sheet_ => sh_);
            wb_.sheets(sh_).rows(curr_row_-1)(curr_col_+col_).style := title_xfId_;
         END IF;
      END IF;
      IF col_headers_ THEN
         Cell (
            curr_col_+col_, curr_row_, desc_tab_(col_).col_name, sheet_ => sh_,
            fontId_ => hdr_font_, fillId_ => hdr_fill_
         );
      END IF;
      CASE
         -- Codes for various forms of number (float, number, binary_double)
         WHEN desc_tab_(col_).col_type IN (2, 100, 101) THEN
            dbms_sql.define_array (crs_, col_, n_tab_, bulk_sz_, 1);
         -- Codes for DATE + TIMESTAMP types (with and without time-zone detail)
         WHEN desc_tab_(col_).col_type IN (12, 178, 179, 180, 181, 231) THEN
            dbms_sql.define_array (crs_, col_, d_tab_, bulk_sz_, 1);
            IF not col_numFmts_.exists(col_) THEN
               col_numFmts_(col_) := Get_numFmt(dft_fmt_date_short_);
            END IF;
         -- Codes for CHAR + VARCHAR types
         WHEN desc_tab_(col_).col_type IN (1, 8, 9, 96, 112) THEN
            dbms_sql.define_array (crs_, col_, v_tab_, bulk_sz_, 1);
         -- Other stuff (like BLOBs) we can't easily encode into Excel, so we ignore!
         ELSE
            null;
      END CASE;
      widths_(col_) := 8;
   END LOOP;
   curr_row_ := curr_row_ + CASE WHEN col_headers_ THEN 1 ELSE 0 END;

   row_count_ := 0;
   LOOP -- loop for each "chunk" of rows fetched
      rows_fetched_ := dbms_sql.fetch_rows(crs_);
      IF rows_fetched_ = 0 THEN goto no_rows_in_chunk; END IF;
      row_count_    := row_count_ + rows_fetched_;
      FOR col_ IN 1 .. col_count_ LOOP
         CASE
            WHEN desc_tab_(col_).col_type IN (2, 100, 101) THEN
               -- Numbers
               Dbms_Sql.Column_Value (crs_, col_, n_tab_);
               FOR i_ IN 0 .. rows_fetched_ - 1 LOOP
                  IF n_tab_(i_+n_tab_.first) IS NOT null THEN
                     Cell (
                        curr_col_+col_, curr_row_+i_, value_ => n_tab_(i_+n_tab_.first),
                        numFmtId_ => CASE WHEN col_numFmts_.exists(col_) THEN col_numFmts_(col_) END,
                        sheet_    => sh_
                     );
                  ELSE
                     CellB (
                        curr_col_+col_, curr_row_+i_, null, sheet_ => sh_,
                        numFmtId_ => CASE WHEN col_numFmts_.exists(col_) THEN col_numFmts_(col_) END
                     );
                  END IF;
               END LOOP;
               n_tab_.delete;
            WHEN desc_tab_(col_).col_type IN (12, 178, 179, 180, 181, 231) THEN
               -- Dates
               Dbms_Sql.Column_Value(crs_, col_, d_tab_);
               FOR i_ IN 0 .. rows_fetched_ - 1 LOOP
                  IF d_tab_(i_+d_tab_.first) IS NOT null THEN
                     Cell (
                        curr_col_+col_, curr_row_+i_, value_ => d_tab_(i_+d_tab_.first),
                        numFmtId_ => CASE WHEN col_numFmts_.exists(col_) THEN col_numFmts_(col_) END,
                        sheet_    => sh_
                     );
                     widths_(col_) := 12; -- assumes dd/mm/yyyy
                  ELSE
                     CellB (
                        curr_col_+col_, curr_row_+i_, null, sheet_ => sh_,
                        numFmtId_ => CASE WHEN col_numFmts_.exists(col_) THEN col_numFmts_(col_) END
                     );
                  END IF;
               END LOOP;
               d_tab_.delete;
            WHEN desc_tab_(col_).col_type IN (1, 8, 9, 96, 112) THEN
               -- Text
               Dbms_Sql.Column_Value (crs_, col_, v_tab_);
               FOR i_ IN 0 .. rows_fetched_-1 LOOP
                  IF v_tab_(i_+v_tab_.first) IS NOT null THEN
                     Cell (curr_col_+col_, curr_row_+i_, value_str_ => v_tab_(i_+v_tab_.first), sheet_ => sh_);
                     data_len_ := length(v_tab_(i_+v_tab_.first));
                     widths_(col_) := least (greatest(widths_(col_),data_len_), 60);
                  ELSE
                     CellB (
                        curr_col_+col_, curr_row_+i_, null, sheet_ => sh_,
                        numFmtId_ => CASE WHEN col_numFmts_.exists(col_) THEN col_numFmts_(col_) END
                     );
                  END IF;
               END LOOP;
               v_tab_.delete;
         END CASE;
      END LOOP;
      << no_rows_in_chunk >>
      EXIT WHEN rows_fetched_ != bulk_sz_;
      curr_row_ := curr_row_ + rows_fetched_;
   END LOOP; -- loop for each column in the result set

   -- set column widths
   ix_ := widths_.first + col_pos_ - 1;
   WHILE ix_ IS not null LOOP
      Set_Column_Width (ix_, widths_(ix_), sh_);
      ix_ := widths_.next(ix_);
   END LOOP;

   Dbms_Sql.Close_Cursor (crs_);

EXCEPTION
   WHEN others THEN
      IF dbms_sql.is_open (crs_) THEN
         dbms_sql.close_cursor (crs_);
      END IF;
END Query2Sheet;

PROCEDURE Do_Binding (
   crs_   IN OUT INTEGER,
   binds_ IN OUT NOCOPY bind_arr )
IS
   bind_id_ VARCHAR2(50) := binds_.first;
BEGIN
   LOOP
      EXIT WHEN bind_id_ IS null;
      CASE binds_(bind_id_).datatype
         WHEN 'STRING' THEN Dbms_Sql.Bind_Variable (crs_, bind_id_, binds_(bind_id_).s_val);
         WHEN 'NUMBER' THEN Dbms_Sql.Bind_Variable (crs_, bind_id_, binds_(bind_id_).n_val);
         WHEN 'DATE'   THEN Dbms_Sql.Bind_Variable (crs_, bind_id_, binds_(bind_id_).d_val);
      END CASE;
      bind_id_ := binds_.next(bind_id_);
   END LOOP;
END Do_Binding;

-- Query2Sheet() => Using SQL, with binding
PROCEDURE Query2Sheet (
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_headers_ IN BOOLEAN        := true,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   crs_   INTEGER := Dbms_Sql.Open_Cursor;
   throw_ INTEGER;
BEGIN
   Dbms_Sql.Parse (crs_, sql_, dbms_sql.native);
   Do_Binding (crs_, binds_);
   throw_ := Dbms_Sql.Execute(crs_); -- ignore
   Query2Sheet (
      col_count_, row_count_, crs_, col_headers_, col_pos_, row_pos_, sheet_,
      title_, title_xfId_, hdr_font_, hdr_fill_, col_fmts_
   );
END Query2Sheet;

-- Query2Sheet() => Using SQL, no binding
PROCEDURE Query2Sheet (
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   col_headers_ IN BOOLEAN        := true,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   binds_ bind_arr := bind_arr();
BEGIN
   Query2Sheet (
      col_count_, row_count_, sql_, binds_, col_headers_, col_pos_, row_pos_, sheet_,
      title_, title_xfId_, hdr_font_, hdr_fill_, col_fmts_
   );
END Query2Sheet;

-- Query2Sheet() => Using REFCURSOR
PROCEDURE Query2Sheet (
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   rc_          IN OUT NOCOPY SYS_REFCURSOR,
   col_headers_ IN BOOLEAN        := true,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   crs_ INTEGER := dbms_sql.to_cursor_number (rc_);
BEGIN
   Query2Sheet (
      col_count_, row_count_, crs_, col_headers_, col_pos_, row_pos_, sheet_,
      title_, title_xfId_, hdr_font_, hdr_fill_, col_fmts_
   );
END Query2Sheet;

PROCEDURE Query2SheetAndAutofilter ( -- with Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   shift_ PLS_INTEGER := CASE WHEN title_ IS NOT null THEN 1 ELSE 0 END;
BEGIN
   Query2Sheet (
      col_count_   => col_count_,
      row_count_   => row_count_,
      sql_         => sql_,
      binds_       => binds_,
      col_headers_ => true,
      col_pos_     => col_pos_,
      row_pos_     => row_pos_,
      sheet_       => sheet_,
      title_       => title_,
      title_xfId_  => title_xfId_,
      hdr_font_    => hdr_font_,
      hdr_fill_    => hdr_fill_,
      col_fmts_    => col_fmts_
   );
   Set_Autofilter (
      col_pos_, col_pos_ + col_count_ - 1,
      row_pos_ + shift_, row_pos_ + row_count_ + shift_, sheet_
   );
END Query2SheetAndAutofilter;

PROCEDURE Query2SheetAndAutofilter ( -- no Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   binds_ bind_arr := bind_arr();
BEGIN
   Query2SheetAndAutofilter (
      col_count_, row_count_, sql_, binds_, col_pos_, row_pos_, sheet_,
      title_, title_xfId_, hdr_font_, hdr_fill_, col_fmts_
   );
END Query2SheetAndAutofilter;

PROCEDURE Query2SheetAndAutofilter ( -- ref-cursor
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   rc_          IN OUT NOCOPY SYS_REFCURSOR,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   hdr_font_    IN PLS_INTEGER    := null,
   hdr_fill_    IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   crs_   INTEGER     := dbms_sql.to_cursor_number (rc_);
   shift_ PLS_INTEGER := CASE WHEN title_ IS NOT null THEN 1 ELSE 0 END;
BEGIN
   Query2Sheet (
      col_count_, row_count_, crs_, true, col_pos_, row_pos_, sheet_,
      title_, title_xfId_, hdr_font_, hdr_fill_, col_fmts_
   );
   Set_Autofilter (
      col_pos_, col_pos_ + col_count_ - 1,
      row_pos_ + shift_, row_pos_ + row_count_ + shift_, sheet_
   );
END Query2SheetAndAutofilter;

PROCEDURE Query2Table ( -- with Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   binds_       IN OUT NOCOPY bind_arr,
   table_style_ IN VARCHAR2,
   tbl_name_    IN VARCHAR2       := null,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   shift_ PLS_INTEGER := CASE WHEN title_ IS NOT null THEN 1 ELSE 0 END;
BEGIN
   Query2Sheet (
      col_count_, row_count_, sql_, binds_, true, col_pos_, row_pos_, sheet_,
      title_, title_xfId_, col_fmts_ => col_fmts_
   );
   Set_Table (
      col_pos_, col_pos_ + col_count_ - 1,
      row_pos_ + shift_, row_pos_ + row_count_ + shift_,
      table_style_, tbl_name_, sheet_
   );
END Query2Table;

PROCEDURE Query2Table ( -- no Binds
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   sql_         IN VARCHAR2,
   table_style_ IN VARCHAR2,
   tbl_name_    IN VARCHAR2       := null,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   binds_ bind_arr := bind_arr();
BEGIN
   Query2Table (
      col_count_, row_count_, sql_, binds_, table_style_, tbl_name_,
      col_pos_, row_pos_, sheet_, title_, title_xfId_, col_fmts_
   );
END Query2Table;
-------------------------------------------------------------
PROCEDURE Query2Table ( -- ref-cursor
   col_count_   IN OUT NOCOPY PLS_INTEGER,
   row_count_   IN OUT NOCOPY PLS_INTEGER,
   rc_          IN OUT SYS_REFCURSOR,
   table_style_ IN VARCHAR2,
   tbl_name_    IN VARCHAR2       := null,
   col_pos_     IN PLS_INTEGER    := 1,
   row_pos_     IN PLS_INTEGER    := 1,
   sheet_       IN PLS_INTEGER    := null,
   title_       IN VARCHAR2       := null,
   title_xfId_  IN PLS_INTEGER    := null,
   col_fmts_    IN tp_numFmt_cols := tp_numFmt_cols() )
IS
   crs_   INTEGER     := dbms_sql.to_cursor_number (rc_);
   shift_ PLS_INTEGER := CASE WHEN title_ IS NOT null THEN 1 ELSE 0 END;
BEGIN
   Query2Sheet (
      col_count_, row_count_, crs_, true, col_pos_, row_pos_, sheet_,
      title_, title_xfId_, col_fmts_ => col_fmts_
   );
   Set_Table (
      col_pos_, col_pos_ + col_count_ - 1,
      row_pos_ + shift_, row_pos_ + row_count_ + shift_,
      table_style_, tbl_name_, sheet_
   );
END Query2Table;


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
   fonts_('bld_lg')      := Get_Font (bold_ => true, fontsize_ => 14);
   fonts_('bld_wht')     := Get_Font (rgb_ => 'FFFFFFFF', bold_ => true);
   fonts_('bld_dk_bl')   := Get_Font (rgb_ => 'FF244062', bold_ => true);
   fonts_('bld_lt_bl')   := Get_Font (rgb_ => 'FFDCE6F1', bold_ => true);
   fonts_('bld_ltbl_lg') := Get_Font (rgb_ => 'FFDCE6F1', bold_ => true, fontsize_ => 14);
   fonts_('bld_wht_lg')  := Get_Font (rgb_ => 'FFFFFFFF', bold_ => true, fontsize_ => 14);
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
   fills_('vlt_grey')    := Get_Fill ('solid', 'FFF2F2F2');
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

   xf_('dkblhd') := Get_XfId (fontName_ => 'head1', fillName_ => 'dk_blue');
   -- numFmtName_, fontName_, fillName_, borderName_, alignName_

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
   extra_blurb_ IN VARCHAR2,
   show_user_   IN BOOLEAN     := true,
   sheet_       IN PLS_INTEGER := null )
IS
   row_ NUMBER := 2;
   sh_  PLS_INTEGER := nvl(sheet_, wb_.sheets.count);
BEGIN

   -- Information about the report is static, with the only option being as to
   -- whether we show the user who printed the report
   CellS (2, row_, 'Report Information', xfName_ => 'dkblhd', sheet_ => sh_);
   CellS (3, row_, '', xfName_ => 'dkblhd', sheet_ => sh_);
   row_ := row_ + 1;
   CellS (2, row_, 'Report Name', fontName_ => 'bold', sheet_ => sh_);
   CellS (3, row_, value_str_ => report_name_);
   row_ := row_ + 1;
   CellS (2, row_, 'Executed at', fontName_ => 'bold', sheet_ => sh_);
   CellS (3, row_, value_str_ => to_char(sysdate, 'YYYY-MM-DD HH24:MI:SS'), sheet_ => sh_);
   row_ := row_ + 1;
   IF show_user_ THEN
      CellS (2, row_, 'Executed by', fontName_ => 'bold', sheet_ => sh_);
      CellS (3, row_, value_str_ => user, sheet_ => sh_);
      row_ := row_ + 1;
   END IF;

   -- Then we print the parameter headers, with the values output in a loop
   row_ := row_ + 1;
   CellS (2, row_, 'Parameters', xfName_ => 'dkblhd', sheet_ => sh_);
   CellS (3, row_, 'Value', xfName_ => 'dkblhd', sheet_ => sh_);
   CellS (4, row_, 'Additional Info', xfName_ => 'dkblhd', sheet_ => sh_);
   row_ := row_ + 1;
   FOR i_ IN params_.FIRST .. params_.LAST LOOP
      CellS (2, row_, params_(i_).param_name, fontName_ => 'bold', sheet_ => sh_);
      CellS (3, row_, value_str_ => params_(i_).param_value, sheet_ => sh_);
      CellS (4, row_, value_str_ => params_(i_).additional_info, sheet_ => sh_);
      row_ := row_ + 1;
   END LOOP;
   row_ := row_ + 1;

   IF extra_blurb_ IS NOT null THEN
      CellS (2, row_, 'Additional report information', xfName_ => 'dkblhd', sheet_ => sh_);
      CellB (3, row_, xfName_ => 'dkblhd', sheet_ => sh_);
      CellB (4, row_, xfName_ => 'dkblhd', sheet_ => sh_);
      row_ := row_ + 1;
      Mergecells (
         tl_col_ => 2, tl_row_ => row_,
         br_col_ => 4, br_row_ => row_+7, sheet_ => sh_
      );
      CellS (
         col_ => 2, row_ => row_, value_str_ => extra_blurb_,
         alignName_ => 'wrap', fillName_ => 'vlt_grey', sheet_ => sh_
      );
   END IF;

   Set_Column_Width (2, 25, sh_);
   Set_Column_Width (3, 40, sh_);
   Set_Column_Width (4, 40, sh_);

END Create_Params_Sheet;

END Nyce_Xlsx;
/
