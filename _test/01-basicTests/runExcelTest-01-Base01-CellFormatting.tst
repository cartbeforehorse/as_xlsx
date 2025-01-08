PL/SQL Developer Test script 3.0
63
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '01-basicTests';
   test_name_   CONSTANT VARCHAR2(30) := 'NumFormats';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(60);
   xl_blob_     BLOB;
   row_         PLS_INTEGER := 2;
BEGIN

   Nyce_Xlsx.Init_Workbook;
   Nyce_Xlsx.Set_Sheet_Name (1, 'Formatting');
   Nyce_Xlsx.CellS (2, row_, 'Formatting tests ==>');
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, 'bold font:');
   Nyce_Xlsx.CellS (3, row_, 'some bold text', fontName_ => 'bold');
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, 'italic font:');
   Nyce_Xlsx.CellS (3, row_, 'some italic text', fontName_ => 'italic');
   row_ := row_ + 1;

   Nyce_Xlsx.Cells (2, row_, 'bold white font, large, with backkground:');
   Nyce_Xlsx.CellS (3, row_, 'Important Title', fontName_ => 'bld_wht_lg', fillName_ => 'md_dk_blue');
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, 'hyperlink:');
   Nyce_Xlsx.Hyperlink (3, row_, 'https://cartbeforehorse.com', 'click me');
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, 'GBP currency format:');
   Nyce_Xlsx.CellN (3, row_, 123.657, numFmtName_ => 'gbp_curr2');
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, 'Two decimal places:');
   Nyce_Xlsx.CellN (3, row_, 0.56, numFmtName_ => '2dp');
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, '10 decimal places, "General" formatting:');
   Nyce_Xlsx.CellN (3, row_, 43563.9899665367);
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, 'Date to the second:');
   Nyce_Xlsx.CellD (3, row_, to_date('01012024-134621','DDMMYYYY-HH24MISS'), numFmtName_ => 'dthms_mid');
   row_ := row_ + 1;

   Nyce_Xlsx.CellS (2, row_, 'Same date as month/year only:');
   Nyce_Xlsx.CellD (3, row_, to_date('01012024-134621','DDMMYYYY-HH24MISS'), numFmtName_ => 'Mmm yyyy');
   row_ := row_ + 1;

   Nyce_Xlsx.Comment (2, 2, 'If you can find me, I will build it', 'Bob the Builder', 300, 200, 1);

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
7
used_
italic_
bold_
fmt_mask_
md5_hash_
xf_count_
xf_.fontId
