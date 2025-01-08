PL/SQL Developer Test script 3.0
47
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '02-tables';
   test_name_   CONSTANT VARCHAR2(30) := '04-q2sNotTableWithTitle';
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
   Nyce_Xlsx.Query2SheetAndAutofilter (
      col_count_  => col_count_,
      row_count_  => row_count_,
      sql_        => sql_,
      col_pos_    => 3,
      row_pos_    => 3,
      sheet_      => sheet_,
      title_      => 'This table is pretty',
      title_xfId_ => Nyce_Xlsx.Get_XfId (
         alignment_ => Nyce_Xlsx.Get_Alignment (horizontal_ => 'centerContinuous')
      )
   );

   FOR c_ IN 3 .. col_count_ + 3 LOOP
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
0
5
3
range_.tl.c
range_.br.c
rollup_type_
