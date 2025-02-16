# Create an Excel-file with PL/SQL

`Nyce_Xlsx` is an engine written in PL/SQL to generate Excel files directly out of Oracle databases.

The initial version was created by Anton Scheffer and called `AS_XLSX`. Our version is based on a copy taken from his on 2021-10-04.

[His original blog is here >>](https://technology.amis.nl/languages/oracle-plsql/create-an-excel-file-with-plsql/)

## Licensing

Please see the [License file](LICENSE.md) to understand your rights.

# Background, Branching, and Regression Testing

Anton also has a Git repo [here >>](https://github.com/antonscheffer/as_xlsx).  There are seveal reasons that we have not branched directly from his repo (in the classical Git sense) – the most obvious being that we created this repo before he created his!  That said, we are trying to keep up with changes being applied to his, by manually back-porting his functionality to ours.

However, we also believe that our version has several benefits and improvements over his, which is why we decided to "branch" from his version in the first place.

## Reasons for *not* re-merging

In order to implement some of the functionality that we want, we've had to make **significant** changes to the code-base, which makes re-merging the code impractical.

The big fundamental change that we've made in our version is to make use of Oracle's `Dbms_XmlDom` XML engine to build XML files.  This allows for cleaner coding standards, easier debugging and improved version/diff-tracking.  It also reduces the likelihood of introducing XML-formatting bugs, and *probably* leads to performance enhancements since manipulating CLOB strings can be slow in some circumstances (this is a theoretical benefit at the moment though as we haven't underone heavy stress-tested yet - December 2024).

Our version now supports pivot-tables, which his does not, which is a **big** job to merge back into the original!

Other enhancements have focused on improving the API interface, allowing the calling program (i.e. your code) to have a smaller footprint.  We've focused on reducing the syntactic overhead (as compared with the original version).  In particular, when trying to format large areas of an Excel sheet, we found the original version verbose to the point that it starts to drown out the calling program's code footprint, distracting from the business logic, and the overall "purpose" (or "flow") of the program.

## Testing and Learining

Finally, this version includes executable test-scripts and a test-manager file to facilitate regression testing.  Admittedly, these files are IDE specific, but they also allow for a more standardised testing process.  It's nice to know that the most recent changes don't break older functionality!  The test scripts will also be useful for newcomers who want to learn how it the interface works.  *Aside:* the `_test` directory also containssome `.sql` files to create tables and inject data that the Excel files 

# Installation

## The basics

Of course, you can download this code by the usual Git command-line method:

    git clone https://github.com/cartbeforehorse/as_xlsx.git

However, we'd suggest cloning through your favourite Git GUI tool.  Understanding change-control is so much easier whe you have a graphical interface to visualise it all.

The primary package `Nyce_Xlsx`  does have dependencies, that are included in this repo.  In all, you'll need to install 3 PL/SQL packages (in the following order):

 - `Nyce_Utils`
 - `Nyce_Xml`
 - `Nyce_Xlsx`

You can deploy each of these into the database from a SQL*Plus prompt using the `@` command to deploy each package.  For example: `SQL> @nyce_utils.pck` should deploy the first file (assuming you're located in the repository's root directory, of course).  Again though, you may be more comfortable deploying packages through your own IDE.

Also, in order to get the encryption functionality working, you'll need to make sure that your database has Oracle's `Dbms_Crypto` modules installed.  The package will deploy without it, but attempts to encrypt an Excel file will silently fail.

## Branching and Support

We would love to have help improving our code, and to receive new ideas on how to evolve the application/toolset.  So please do feel free to branch from our main branches.

Alternatively, add your issues and improvement ideas to the "Issues" and "Projects" tabs on this GitHub page and we'll see what we can do to help.

---

# Table of Contents

- [Create an Excel-file with PL/SQL](#create-an-excel-file-with-plsql)
  - [Licensing](#licensing)
- [Background, Branching, and Regression Testing](#background-branching-and-regression-testing)
  - [Reasons for *not* re-merging](#reasons-for-not-re-merging)
  - [Testing and Learining](#testing-and-learining)
- [Installation](#installation)
  - [The basics](#the-basics)
  - [Branching and Support](#branching-and-support)
- [Table of Contents](#table-of-contents)
- [Requirements and Standards](#requirements-and-standards)
  - [Oracle Version Support](#oracle-version-support)
  - [Tooling](#tooling)
  - [Coding Standards](#coding-standards)
- [Basic Usage](#basic-usage)
  - [Quick-start: Extracting Data from DB to Sheet](#quick-start-extracting-data-from-db-to-sheet)
  - [Using bind variables](#using-bind-variables)
  - [Using a ref-cursor](#using-a-ref-cursor)
- [Formatting](#formatting)
  - [Understanding Xf](#understanding-xf)
  - [`Get*()` functions\`](#get-functions)
- [Gotchas](#gotchas)
- [New in Version 2](#new-in-version-2)
  - [`fonts_` and `fills_` variables](#fonts_-and-fills_-variables)
  - [Binded SQL statements](#binded-sql-statements)
  - [Query, Autofilter and Format simultaneously](#query-autofilter-and-format-simultaneously)
  - [Report Overview Page](#report-overview-page)
  - [Column auto-width](#column-auto-width)

# Requirements and Standards

## Oracle Version Support
You'll need Oracle Database 19c or greater to use features included in this package.  We tried to make it backwardly compatible to Oracle 12c, but that proved to be too much work (and have too little benefit).  Sorry!

Please feel free to branch though!

## Tooling
This code was developed on an IDE called [PL/SQL Developer](https://www.allroundautomations.com/products/pl-sql-developer/), which is by far the best tool for organising Oracle work (in our humble opinion).  The tool organises PL/SQL package headers and body into a single `pck` file, which makes more sense than the 2-file structure most IDEs use.  `pck` files can still be compiled as a normal SQL file for deployment purposes though, so this shouldn't bother anyone who normally use the 2-file system.

PL/SQL Developer also has a native format for storing test files (and test-manager files).  However, since most developers of PL/SQL code don't bother with formalised testing, I doubt this will bother too many people!  It should be reassuring to know that on-going development is validated against historical/regressive test scenarios, and anyway you'll find that the test-code can be easily extracted from `tst` files should you want to run them in a different IDE.

## Coding Standards
Yeah, okay, this is a bit of a moan.  But if you plan to make changes, please observe the coding conventions used in the package.  In particular, please define variables with a trailing underscore: `_`.  It is a constant source of amusement to us that PL/SQL coders want to define their variables as "variables" and parameters as "parameters" by using `v_`, `p_` or `i_` prefixes (which never seem to be consistent anyway).  The least interesting property of a v_variable or p_parameter is the fact that it is a v_variable or p_parameter!  Every other programming language in the world aspires to make itself more readable by describing variables as the data they store, and deliberately abstracting away from the the fact that they are v_variables (just as English abstracs away from describing w_each w_word w_as w_a w_word).  And yet PL/SQL code generally appears to buck this trend.  Go figure!  Anyhoo, sorry about that little digression.  Funny how some details can get under one's skin, innit!! :-P


# Basic Usage

## Quick-start: Extracting Data from DB to Sheet

The most common requirement for an Excel-database tool (such as `Nyce_Xlsx`) is to transfer data from a database-table into a more manageable Excel sheet.  The simple shortcut method of doing this is as follows:
```
DECLARE
   cols_  PLS_INTEGER;
   rows_  PLS_INTEGER;
BEGIN
   Nyce_Xlsx.query2sheet (cols_, rows_, 'select * from dual');
   Nyce_Xlsx.Save ('MY_DIR', 'my.xlsx');
END;
```

This will turn your SQL statement into a 2D data-array in an MS Excel sheet, and then save that excel sheet onto your file-system.  The directory `MY_DIR` needs to be defined as an oracle directory (which is subsequently mapped to a physical directory, of course).  The filename is defined in the second parameter.

No formatting is added by default.  Formatting is quite a large subject which we'll tackle a little later.

Note that the procedure passes back `row_count_` and `col_count_` values in `OUT` variables which is useful because we're unlikely to know the number of records in our dataset before run-time.  The information allows us to continue writing data to the sheet below (or to the right) of the populated table.

There are equivalent functions named `Query2SheetAndAutofilter()` and `Query2Table()` which add filters to the completed dataset, or put the data in an Excel table (which can be colour-formatted to our preference).

## Using bind variables
If your SQL depends on data that's only available at run-time, then you could always write a solution to dynamically build a SQL statement at run-time.  Howver, it's a little more proffessional (and easier to debug) if you keep your SQL as rigid as possible and instead include placeholders to support your dynamic data.  You can use the `Bind_Value()` overloaded function to declare the values that should be binded to your SQL; note that it accepts `VARCHAR2`, `NUMBER` and `DATE` datatypes, whose type will automatically be recognised during execution.

In this example, we also place the resulting dataset to start on cell `C3` (column 3, row 3).  If not supplied, these values will default to cell A1 (otherwise referred to as [1, 1]).

```
DECLARE
   cols_  PLS_INTEGER;
   rows_  PLS_INTEGER;
   sheet_ PLS_INTEGER := Nyce_Xlsx.New_Sheet ('Data and Tables');
   binds_ nyce_xlsx.bind_arr,
   sql_  VARCHAR2(2000) := q'[
      SELECT e.company "Company", e.identity_type "Identity Type", e.identity "Identity",
             e.category "Category", e.currency "Currency", e.amount "Amount", e.tax "Tax"
      FROM   entities5dim_tab e
      WHERE  e.company = :comp
        AND  e.amount  > :amt
   ]';
BEGIN
   Nyce_Xlsx.Bind_Value (binds_, 'comp', 'HLD');
   Nyce_Xlsx.Bind_Value (binds_, 'amt', 500);
   Nyce_Xlsx.Query2SheetAndAutofilter (
      col_count_  => cols_,
      row_count_  => rows_,
      sql_        => sql_,
      binds_      => binds_,
      col_pos_    => 3,
      row_pos_    => 3,
      sheet_      => sheet_,
      title_      => 'This table is pretty',
      title_xfId_ => Nyce_Xlsx.Get_XfId (
         alignment_ => Nyce_Xlsx.Get_Alignment (horizontal_ => 'centerContinuous')
      )
   );
   Nyce_Xlsx.Save ('MY_DIR', 'my.xlsx');
END;
```

Note also the slightly uncomfortable (and very verbose) use of a nested function to obtain an `Xf` styling value.  We'll have a look at how this works a little later.

## Using a ref-cursor

The same result can also be achieved through the use of `REFCURSOR`s.  You'll also notice that in this example, we make use of the `col_fmts_` parameter, allowing us to individually format numeric and date type columns to our liking.  Since column 6 in our SQL statement is an `"Amount"` value, we opt to set it as a 2-decimal, thousand-separated number.  We'll look in more detail at how formatting works a little later.

As you proably know, date-types in Excel are really nothing more than number values which have a special format mask.  All of the `Query2*()` functions discussed here will automatically detect date-types, and format them with the predefined short-date format (unless they were already overriden by the `fmts_` variable discussed in the preceding paragraph).  At the start of each execution-session, the short date format is set to `yyyy-mm-dd` (going with the msb programming principle!!).  But of course, you can change the default to comply with your client's regionalised requirements.  The following example used `Set_Dft_Fmt_Date_Short()` to set the European date format for the remainder of the program.

```
DECLARE
   cols_  PLS_INTEGER;
   rows_  PLS_INTEGER;
   sheet_ PLS_INTEGER  := Nyce_Xlsx.New_Sheet ('Data and Tables');
   comp_  VARCHAR2(50) := 'HLD';
   amt_   NUMBER       := 500;
   qry_   sys_refcursor;
   fmts_  nyce_xlsx.tp_numFmt_cols;
BEGIN
   -- Short date format defaults to "yyyy-mm-dd"; changing this value affects date-columns in `Query2Table()`
   Nyce_Xlsx.Set_Dft_Fmt_Date_Short ('dd-mm-yyyy');
   -- Unless you declare a format-mask on colunm 8 that includes a time element, only the date element of the
   -- current time will be visible in the output Excel sheet.  The actual value in Excel will of course still
   -- include the time element, but the format-mask truncates what the user can see.
   fmts_(6) := Nyce_Xlsx.Get_NumFmt ('#,##0.00'); -- sets the format-mask for the "Amount" column data
   OPEN qry_ FOR
      SELECT e.company "Company", e.identity_type "Identity Type", e.identity "Identity",
             e.category "Category", e.currency "Currency", e.amount "Amount", e.tax "Tax",
             sysdate "Current Time"
      FROM   entities5dim_tab e
      WHERE  e.company = comp_
        AND  e.amount  > amt_;
   Nyce_Xlsx.Query2Table (
      col_count_   => cols_,
      row_count_   => rows_,
      rc_          => qry_,
      table_style_ => 'TableStyleLight21', -- sets the colour/style of your table
      tbl_name_    => 'MyCleverTable',
      col_fmts_    => fmts_
   );
   Nyce_Xlsx.Set_Column_Width (5, 25);
   Nyce_Xlsx.Save ('MY_DIR', 'my.xlsx');
END;
```

On a final note, you can also see that we're able to define the width of each column with `Set_Column_Width()` to better and display our data.


# Formatting

## Understanding Xf

We've all formatted a cell in Excel before, but less do we worry about the breadth of options lying behind it all.  Fonts, colours, borders, shading, patterns... they all need to be selected by the user as a means of improved presentation, and all these settings must then be stored in the sheet.

`Nyce_Xlsx` gives the calling program control over the following properties of a cell, and as you see, each property is further broken down into sub-properties
 - **Font:** font-face, font-size, colour, boldness, italicness, underline
 - **Shading/Fill:** pattern, colour-1, colour-2
 - **Number-format:** number-format (includingdate-format)
 - **Border:** border-top, border-right, border-bottom, border-left
 - **Alignment:** vertical-alignment, horizontal-alignment, word-wrap

Each one of the above parameters can be changed independently for each cell on every sheet of your output Excel document.  [`Xf`](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/4fe94af3-cd05-427a-b32d-1f27a85c412a) is the name given by Excel to refer to a given combination of all the above properties.  In other words, if you have an `xfId`, you can work backwards to find out what font, what colour and what border combination that `xfId` represents.  A cell with that same `xfId` will display with all of those formatting properties.

Breaking the prolem down to the next level, it should come as no surprise to learn that each font configuration has a conrresponding `fontId_`, each shading pattern has a `fillId_`, and each border pattern has a `borderId_` (etc., etc.).  So another way of describing `Xf` is that it is the top level of a combination of properties, which themselves represent a deeper combination of more specialised properties.

Of course, each `xfId_` can only represent one combination of properties.  But perhaps less obviously, each combination of properties should only ever be represented by a single `xfId_.`.

## `Get*()` functions`

For each property in the above list, `Nyce_Xlsx` provides a `Get*()` function.  Let's start with `Get_Font()`.

The concept is that you need to define your formatting yourself using the function `Get_Font()`.  The `Get` prefix is a little misleading to be honest, in the sense that it doesn't necessarily "get" anything at all.  What it does is to search for the combination of format-settings that you've given it, and check whether or not that combination exists.  If the combination already exists, then the existing ID already assigned to the combination is returned; otherwise a new ID is generated and handed back.

Example usage of the `Get` functions:

... to be continued....................................................
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


