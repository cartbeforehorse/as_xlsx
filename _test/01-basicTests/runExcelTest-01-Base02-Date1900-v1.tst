PL/SQL Developer Test script 3.0
40
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '01-basicTests';
   test_name_   CONSTANT VARCHAR2(30) := 'Date1900-v1';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(80);
   xl_blob_     BLOB;
   sheet_       PLS_INTEGER;
   base_date_   DATE := to_date ('1900-01-01','YYYY-MM-DD');
   feb28_       DATE := to_date ('1900-02-28','YYYY-MM-DD');
   --feb29_       DATE := to_date ('1900-02-29','YYYY-MM-DD');
   -- excel thinks there's a 29/02/1900, Oracle bugs out if you try to declare this date.
   -- Go see how the above data is displayed in the output Excel file for those dates above!!
   mar1_        DATE := to_date ('1900-03-01','YYYY-MM-DD');
   mar2_        DATE := to_date ('1900-03-02','YYYY-MM-DD');
   test_date_   DATE := to_date ('2024-01-01 14:23:56', 'YYYY-MM-DD HH24:MI:SS');
BEGIN

   Nyce_Xlsx.Init_Workbook;
   Nyce_Xlsx.Set_Sheet_Name (1, 'Date 1900');
   Nyce_Xlsx.CellD (2, 2, base_date_);
   Nyce_Xlsx.CellD (2, 3, base_date_, numFmtName_ => 'dthm_mid', fontName_ => 'bold');
   Nyce_Xlsx.CellD (2, 4, test_date_);
   Nyce_Xlsx.CellD (2, 5, test_date_, numFmtName_ => 'dthms_mid', fontName_ => 'bold');

   Nyce_Xlsx.CellD (2, 7, feb28_);
   --Nyce_Xlsx.CellD (2, 7.5, feb29_);
   Nyce_Xlsx.CellD (2, 8, mar1_);
   Nyce_Xlsx.CellD (2, 9, mar2_);

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
no
5
4
fmt_mask_
num_fmt_id_
numFmtId_
