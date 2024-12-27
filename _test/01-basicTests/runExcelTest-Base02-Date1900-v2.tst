PL/SQL Developer Test script 3.0
22
DECLARE
   qry_    SYS_REFCURSOR;
   width_  PLS_INTEGER;
   height_ PLS_INTEGER;
BEGIN
   OPEN qry_ FOR
      SELECT date '1900-02-26' + level "Secret Date",
             to_char( date '1900-02-26' + level, 'yyyy mon dd' ) "Secret String"
      FROM   dual
      CONNECT BY level < 8;
      Nyce_Xlsx.Clear_Workbook;
      Nyce_Xlsx.Query2SheetAndAutofilter (
         rc_          => qry_,
         directory_   => 'EXCEL_OUT',
         filename_    => 'nyce_xlsx.xlsx',
         useXf_       => true
      );
      Nyce_Xlsx.Set_Column_Width (1, 15);
      Nyce_Xlsx.Set_Column_Width (2, 15);
      -- make sure you have set as_xlsx.use_dbms_crypto = true; in the package specification
      --Ny.save( 'EXCEL_OUT', 'my.xlsx', 'demo' );
END;
0
6
col_
useXf_
format_mask_
fmt_id_
numFmtId_
g_useXf_
