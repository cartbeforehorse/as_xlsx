PL/SQL Developer Test script 3.0
29
DECLARE
   file_name_ CONSTANT VARCHAR2(80) := Cbh_Utils_API.Rep ('TestDate1900_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   sheet_     PLS_INTEGER;
   base_date_ DATE := to_date ('1900-01-01','YYYY-MM-DD');
   feb28_     DATE := to_date ('1900-02-28','YYYY-MM-DD');
   --feb29_     DATE := to_date ('1900-02-29','YYYY-MM-DD');
   -- excel thinks there's a 29/02/1900, Oracle bugs out if you try to declare this date.
   -- Note the output from Excel for the dates above!!
   mar1_      DATE := to_date ('1900-03-01','YYYY-MM-DD');
   mar2_      DATE := to_date ('1900-03-02','YYYY-MM-DD');
   test_date_ DATE := to_date ('2024-01-01 14:23:56', 'YYYY-MM-DD HH24:MI:SS');
BEGIN

   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Date 1900');
   As_Xlsx.CellD (2, 2, base_date_);
   As_Xlsx.CellD (2, 3, base_date_, numFmtId_ => 'dthm_mid', fontId_ => 'bold');
   As_Xlsx.CellD (2, 4, test_date_);
   As_Xlsx.CellD (2, 5, test_date_, numFmtId_ => 'dthms_mid', fontId_ => 'bold');

   As_Xlsx.CellD (2, 7, feb28_);
   --As_Xlsx.CellD (2, 7.5, feb29_);
   As_Xlsx.CellD (2, 8, mar1_);
   As_Xlsx.CellD (2, 9, mar2_);

   As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
   Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

END;
0
4
fmt_mask_
num_fmt_id_
numFmtId_
