PL/SQL Developer Test script 3.0
75
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '02-tables';
   test_name_   CONSTANT VARCHAR2(30) := '05-q2tMultiplesTesting';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(60);
   xl_blob_     BLOB;

   col_count_   PLS_INTEGER;
   row_count_   PLS_INTEGER;
   sheet_       PLS_INTEGER;

   sql_  VARCHAR2(2000) := q'[
      SELECT e.company "Company", e.identity_type "Identity Type", e.identity "Identity",
             e.category "Category", e.currency "Currency", e.amount "Amount", e.tax "Tax"
      FROM   entities5dim_tab e
   ]';
BEGIN

   sheet_ := Nyce_Xlsx.New_Sheet ('Data and Tables');
   Nyce_Xlsx.Query2Table (
      col_count_   => col_count_,
      row_count_   => row_count_,
      sql_         => sql_,
      table_style_ => 'TableStyleMedium28',
      col_pos_     => 3,
      row_pos_     => 3,
      sheet_       => sheet_
   );
-- TableStyleLight1;TableStyleLight21;TableStyleMedium1;TableStyleMedium28;TableStyleDark1;TableStyleDark11
   Nyce_Xlsx.Query2Table (
      col_count_   => col_count_,
      row_count_   => row_count_,
      sql_         => sql_,
      table_style_ => 'TableStyleLight1',
      tbl_name_    => 'Jack',
      col_pos_     => 11,
      row_pos_     => 3,
      sheet_       => sheet_,
      title_       => 'This is table number 2',
      title_xfId_  => Nyce_Xlsx.Get_XfId (
         fontId_    => Nyce_Xlsx.Get_Font (bold_ => true),
         alignment_ => Nyce_Xlsx.Get_Alignment (horizontal_ => 'centerContinuous')
      )
   );

   Nyce_Xlsx.Query2Table (
      col_count_   => col_count_,
      row_count_   => row_count_,
      sql_         => sql_,
      table_style_ => 'TableStyleMedium1',
      col_pos_     => 8,
      row_pos_     => 3 + row_count_ + 3,
      sheet_       => sheet_,
      title_       => 'This table is number 3',
      title_xfId_  => Nyce_Xlsx.Get_XfId (
         fontId_    => Nyce_Xlsx.Get_Font (bold_ => true, italic_ => true),
         alignment_ => Nyce_Xlsx.Get_Alignment (horizontal_ => 'right')
      )
   );

   FOR c_ IN 3 .. 14 LOOP
      Nyce_Xlsx.Set_Column_Width (c_, 15);
   END LOOP;

   IF not create_file_ THEN
      xl_blob_ := Nyce_Xlsx.Finish;
      Dbms_Output.Put_Line ('Test script finished: ' || test_level_ || '-' || test_name_);
   ELSE
      file_name_ := test_level_ || '-' || test_name_ || file_end_;
      Nyce_Xlsx.Save ('EXCEL_OUT', file_name_);
      Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');
   END IF;

END;
1
output_file
1
yes
5
3
range_.tl.c
range_.br.c
rollup_type_
