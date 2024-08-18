PL/SQL Developer Test script 3.0
24
DECLARE
   file_end_    CONSTANT VARCHAR2(20) := Cbh_Utils_API.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_start_  CONSTANT VARCHAR2(20) := 'TestOut_';
   file_name_   VARCHAR2(60);
BEGIN

   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Formatting');
   As_Xlsx.CellS (2, 2, 'some bold text', fontId_ => 'bold');
   As_Xlsx.CellS (2, 3, 'some italic text', fontId_ => 'italic');
   As_Xlsx.Hyperlink (2, 4, 'https://cartbeforehorse.com', 'click me');
   As_Xlsx.CellN (2, 5, 123.657, numFmtId_ => 'gbp_curr2');
   As_Xlsx.CellN (2, 6, 0.56, numFmtId_ => '2dp');
   As_Xlsx.CellN (2, 7, 43563.9899665367);
   As_Xlsx.CellD (2, 8, to_date('01012024-1346','DDMMYYYY-HH24MI'), numFmtId_ => 'dthms_mid');
   As_Xlsx.CellD (2, 9, to_date('01012024-1346','DDMMYYYY-HH24MI'), numFmtId_ => 'Mmm yyyy');

   As_Xlsx.Comment (2, 2, 'This is a silly dingle dongle', 'Bob the Builder', 300, 200, 1);

   file_name_ := file_start_ || 'NumFormats' || file_end_;
   As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
   Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

END;
0
7
used_
italic_
bold_
fmt_mask_
md5_hash_
xf_count_
xf_.fontId
