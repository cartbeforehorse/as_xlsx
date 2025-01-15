PL/SQL Developer Test script 3.0
55
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '02-tables';
   test_name_   CONSTANT VARCHAR2(30) := 'AS05-q2sMultiplesTesting';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(60);
   xl_blob_     BLOB;

   col_count_   PLS_INTEGER;
   row_count_   PLS_INTEGER;
   sheet_       PLS_INTEGER := 1;

   sql_  VARCHAR2(2000) := q'[
      SELECT e.company "Company", e.identity_type "Identity Type", e.identity "Identity",
             e.category "Category", e.currency "Currency", e.amount "Amount", e.tax "Tax"
      FROM   entities5dim_tab e
   ]';
BEGIN

   As_Xlsx.New_Sheet ('Data and Tables');
   As_Xlsx.Query2Sheet (
      sql_, p_sheet => sheet_, p_UseXf => true, p_col => 3, p_row => 3,
      p_table_style => 'TableStyleMedium28'
   );

   As_Xlsx.Query2Sheet (
      sql_, p_sheet => sheet_, p_UseXf => true, p_title => 'This is table number 2',
      p_title_xfid => As_Xlsx.Get_XfId (
         p_fontId    => As_Xlsx.Get_Font('Calibri', p_bold => true),
         p_alignment => As_Xlsx.Get_Alignment(p_horizontal => 'centerContinuous')
      ), p_col => 11, p_row => 3, p_table_style => 'TableStyleLight1'
   );
   As_Xlsx.Query2Sheet (
      sql_, p_sheet => sheet_, p_UseXf => true, p_title => 'This table is number 3',
      p_title_xfid => As_Xlsx.Get_XfId (
         p_fontId    => As_Xlsx.Get_Font('Calibri', p_bold => true),
         p_alignment => As_Xlsx.Get_Alignment(p_horizontal => 'right')
      ), p_col => 8, p_row => 20, p_table_style => 'TableStyleMedium1'
   );


   FOR c_ IN 3 .. 17 LOOP
      As_Xlsx.Set_Column_Width (c_, 15);
   END LOOP;

   IF not create_file_ THEN
      xl_blob_ := As_Xlsx.Finish;
      Dbms_Output.Put_Line ('Test script finished: ' || test_level_ || '-' || test_name_);
   ELSE
      file_name_ := test_level_ || '-' || test_name_ || file_end_;
      As_Xlsx.Save ('EXCEL_OUT', file_name_);
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
