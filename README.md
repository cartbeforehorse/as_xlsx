# Create an Excel-file with PL/SQL

Initial version created by Anton Scheffer
Taken on 2021-10-04

[Please visit his website here >>](https://technology.amis.nl/languages/oracle-plsql/create-an-excel-file-with-plsql/)

## Version Tracking

Why he hasn't created a Git repo himself isn't clear, but here's hoping that this can built
upon by the community.

# Usage

## Extracting Data
```
DECLARE
   cols_  PLS_INTEGER;
   rows_  PLS_INTEGER;
BEGIN
   as_xlsx.query2sheet (cols_, rows_, 'select * from dual');
   as_xlsx.save ('MY_DIR', 'my.xlsx');
END;
```

The above code will turn your SQL statement into a 2D data-array in an MS Excel sheet.
The directory `MY_DIR` needs to be defined as an oracle directory (which is subsequently
mapped to a physical directory, of course).  The filename is defined in the second parameter.

No formatting is added by default (though options have been added in v2 to make your output
prettier).

Note that the procedure passes back a `row_count_` and `col_count_` values in `OUT` variables
which are often useful if your SQL is built dynamically, or you want to locate particular
cells within your grid to format them.


## Formatting

The concept is that you need to define your formatting yourself using functions such as `Get_Font()`
and `Get_Fill()`.  The `Get` prefix to these functions is a little misleading, to be honest, in the
sense that it doesn't "get" anything at all.

Example usage of the `Get` functions:

```
DECLARE
   font_head1_   PLS_INTEGER := as_xlsx.Get_Font (p_rgb=>'FFDBE5F1', p_bold=>true);
   font_bld_     PLS_INTEGER := as_xlsx.Get_Font (p_bold=>true);
   font_bld_wht_ PLS_INTEGER := as_xlsx.Get_Font (p_rgb=>'FFFFFFFF', p_bold=>true);
   font_it_sm_   PLS_INTEGER := as_xlsx.Get_Font (p_italic=>true, p_fontsize=>9);
   bkg_dk_blue_  PLS_INTEGER := as_xlsx.Get_Fill ('solid', 'FF17375D');
   bkg_dk_red_   PLS_INTEGER := as_xlsx.Get_Fill ('solid', 'FF953735');
   -- etc.
```

The above code creates font styles and cell background/fills.  The font-face will default to that
defined in your Excel template (which is normally "Calibri").
 - `font_head_1` is a pinkish font with a bold face
 - `font_bld_` is a black font with a bold face
 - `font_bld_wht_` is white and bold
 - `bkg_dk_red_` defines a solid-background of a dark-red hue

You get the idea.  Anyway, the point is that as you define each style, it gets stored internally
by the package and gets assigned an ID number that is returned to your variable.  If you try to define the
same style (or font) for a second time, `Get_Font()` is intelligent enough to recognise the duplication
and returns the ID of the first style (inherently avoiding storing the same style multiple times).

We later refer to the styles just created with the ID number that was returned to us.  For example,
the following code enters the text `Report Name` in column 2, row 3 (cell B3).  The font will be
black and bold, and the cell will be coloured in a solid red:

```
BEGIN
   -- blah
   as_xlsx.Cell (2, 3, 'Report Name', p_fontId=>font_bld_, p_fillId=>bkg_dk_red_);
   -- blah...
END;
```

# Gotchas

Note that before you start to define your styles, you **have to** initialise your new Excel sheet
and add a new tab to it (a newly created Excel sheet doesn't even have a tab associated with it!).
You can do that with the following code:

```
as_xlsx.Clear_Workbook;
as_xlsx.New_Sheet ('Name of Sheet 1');
-- the above MUST be done before starting to define styles...
font_head1_   PLS_INTEGER := as_xlsx.Get_Font (p_rgb=>'FFDBE5F1', p_bold=>true);
```

If you define styles before creating a sheet on your Excel document, it'll come out all corrupt :-(

A useful little hack introduced in v2 was to add the necessary "setup" functions in the `BEGIN` section
of the package, saving the caller (i.e. you), from having to do so in your packages:

```
CREATE OR REPLACE PACKAGE BODY as_xlsx IS
   -- blah blah...
BEGIN
   Clear_Workbook;
   New_Sheet ('Sheet 1');
END as_xlsx;
```


# New in Version 2


## `fonts_` and `fills_` variables

Nicely formatted Excel sheets can require tens of font styles/colours/background each
of which need to be stored in a variable, adding significant code-clutter in the calling
package.

Though this is a minor addition, the ability to centralise the store of fonts and fills
in a single variable reduces the codebase of the calling package substantially.  We now
recommend that your fonts and fills be stored in the globally accessible `fonts_` and
`fills_` variables.

The package now also initialises a few simple commonly used styles such as a plain bold font,
as well as some coloured backgrounds.  This should save you from needing to set up your own
styles which would clutter up your code (distracting from the real goal you're trying to achieve):

```
PROCEDURE Init_Fonts_And_Fills
IS BEGIN
   fonts_('head1')   := as_xlsx.Get_Font (rgb_ => 'FFDBE5F1', bold_ => true);
   fonts_('bold')    := as_xlsx.Get_Font (bold_ => true);
   fonts_('bld_wht') := as_xlsx.Get_Font (rgb_ => 'FFFFFFFF', bold_ => true);
   fills_('dk_blue') := as_xlsx.Get_Fill ('solid', 'FF17375D');
   fills_('dk_red')  := as_xlsx.Get_Fill ('solid', 'FF953735');
END Init_Fonts_And_Fills;
```

And of course, it's still easy enough to add your own:

```
-- in your own code (or override the existing):
-- ...
   as_xlsx.fonts_('red') := as_xlsx.Get_Font (rgb_ => 'FFFF0000');
-- ...
```


## Binded SQL statements

A public `bind_arr` type has been added to the package.  Binding values into your SQL
saves a ton of string-manipulation mumbo-jumbo in complex scenarios, so this should be
a welcome addition

```
DECLARE
   cust_grp_  VARCHAR2(50) := '12345';
   binds_     as_xlsx.bind_arr;
   cols_      NUMBER;
   rows_      NUMBER;
BEGIN
   binds_(':cust_grp') := as_xlsx.data_binder ('STRING', cust_grp_, null, null);
   binds_(':billed')   := as_xlsx.data_binder ('NUMBER', null, 10000, null);
   as_xlsx.query2sheet (
      col_count_ => cols_,
      row_count_ => rows_,
      sql_       => '
         select c.cust_id, c.cust_name, c.billed_last_year
         from   customer_t c
         where  c.cust_grp = :cust_grp
           and  c.billed_last_year > :billed',
      binds_     => binds_
   );
END;
```

## Query, Autofilter and Format simultaneously

We can now auto-filter the data returned from an SQL statement, in a single call.  By
specifying a pre-defined background-color and font, we can also define how we want to
format the header column of the generated Excel sheet, all in one call.

You still have to format the data-grid yourself though :-(

```
DECLARE
   cust_grp_  VARCHAR2(50) := '12345';
   binds_     as_xlsx.bind_arr;
   cols_      NUMBER;
   rows_      NUMBER;
BEGIN
   binds_(':cust_grp') := as_xlsx.data_binder ('STRING', cust_grp_, null, null);
   binds_(':billed')   := as_xlsx.data_binder ('NUMBER', null, 10000, null);
   as_xlsx.Query2SheetAndAutofilter (
      sql_       => '
         select c.cust_id, c.cust_name, c.billed_last_year
         from   customer_t c
         where  c.cust_grp = :cust_grp
           and  c.billed_last_year > :billed',
      binds_     => binds_,
      UserXf_    => true,
      hdr_font_  => as_xlsx.fonts_('bld_wht'),
      hdr_fill_  => as_xlsx.fills_('dk_blue')
   );
END;
```


## Report Overview Page

If you're generating on-demand reports from the user, it can be useful to include the
date, time and parameters that the user passed to you.  This allows your user to archive
the resulting excel sheet and they will always have a reference to "when" and "what" they
asked for.

```
DECLARE
   binds_     as_xlsx.bind_arr;
BEGIN

   -- set up your bind-variables as shown above; use the same on the parameters sheet
   -- as you did for your actual SQL binding if you want
   binds_(':cust_grp') := as_xlsx.data_binder ('STRING', cust_grp_, null, null);

   As_Xlsx.Create_Params_Sheet (
      report_name_ => 'Customer Order List',
      params_      => params_,
      show_user_   => false, -- option to print the Oracle user's details on the report for their own reference
      sheet_       => 1      -- it would normally be on the front page!
   );

END;
```

## Column auto-width

When a SQL statement is processed to generate a sheet, the columns will now try to
size themselves automatically to suit the size of your data.  They will expand to a
(hard-coded) maximum width of 60 characters.


