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
BEGIN
    as_xlsx.query2sheet ('select * from dual');
    as_xlsx.save ('MY_DIR', 'my.xlsx');
END;
```

The above code will turn your SQL statement into a 2D data-array in an MS Excel sheet.
The directory `MY_DIR` needs to be defined as an oracle directory (mapped to a physical
directory inthe Oracle setup).  The filename is defined in the second parameter.

No formatting is added by default.

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

The above code creates font styles, assuming that the default face is as per your Excel template
(which is normally "Calibri").
 - `font_head_1` is a pinkish font with a bold face
 - `font_bld_` is a b (black) font with nothing font with a bold face
 - `font_bld_wht_` is white and bold
 - `bkg_dk_red_` defines a solid-background dark-red color

You get the idea.  Anyway, the point is that as you define each style, it gets stored internally
by the package, and can be retrieved with its Id value (which is stored in the allocated variable).

As such, we can later refer to the styles with the appropriate IDs as follows:

```
BEGIN
    -- blah
    as_xlsx.Cell (2, 3, 'Report Name', p_fontId=>font_bld_, p_fillId=>bkg_dk_red_);
    -- blah...
END;
```

Here (with sheet number 1 being the implied default), we manipulate cell in column 2, row number 3
by entering the text `Report Name` in the cell.  The font and fill of the cell will be given the
styles as defined earlier.


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

