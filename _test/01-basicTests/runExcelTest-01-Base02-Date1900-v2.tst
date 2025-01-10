PL/SQL Developer Test script 3.0
39
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '01-basicTests';
   test_name_   CONSTANT VARCHAR2(30) := 'Date1900-v2';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(80);
   xl_blob_     BLOB;
   qry_         SYS_REFCURSOR;
   col_fmts_    nyce_xlsx.tp_numFmt_cols;
   col_count_   PLS_INTEGER;
   row_count_   PLS_INTEGER;
BEGIN
   OPEN qry_ FOR
      SELECT date '1900-02-26'+level "Oracle Date",
             to_char (date '1900-02-26'+level, 'yyyy mon dd') "Ora Date as String",
             57 + level "XL Incremental Date"
      FROM   dual
      CONNECT BY level < 8;
   Nyce_Xlsx.Init_Workbook;
   --col_fmts_(3) := Nyce_Xlsx.Get_numFmt ('dd mmm yyyy');
   col_fmts_(3) := nyce_xlsx.numFmt_('dt_mid');
   Nyce_Xlsx.Query2SheetAndAutofilter (col_count_, row_count_, rc_ => qry_, col_fmts_ => col_fmts_, sheet_ => 1);
   Nyce_Xlsx.Set_Column_Width (1, 15);
   Nyce_Xlsx.Set_Column_Width (2, 15);
   Nyce_Xlsx.Set_Column_Width (3, 15);

   -- Following 2 lines will save the file, password protected, where the password is 'demo'
   -- make sure you have set as_xlsx.use_dbms_crypto = true; in the package specification
   --Nyce_Xlsx.save('EXCEL_OUT', 'my.xlsx', 'demo' );

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
6
col_
useXf_
format_mask_
fmt_id_
numFmtId_
g_useXf_
