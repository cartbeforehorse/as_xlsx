PL/SQL Developer Test script 3.0
72
DECLARE
   file_end_    CONSTANT VARCHAR2(20) := Cbh_Utils_API.Rep ('_:P1.zip', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_start_  CONSTANT VARCHAR2(20) := 'TestOut_';
   file_name_   VARCHAR2(60);
   sheet_       PLS_INTEGER := 1;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 3;
   init_row_    PLS_INTEGER := row_;
   data_range_  as_xlsx.tp_cell_range;
   rollup_cols_ as_xlsx.rollup_columns;
   blob_        BLOB;

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

BEGIN

   -- Image File
   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Data and Pivot');

   As_Xlsx.CellS (col_,   row_, 'Identity Type');
   As_Xlsx.CellS (col_+1, row_, 'Identity');
   As_Xlsx.CellS (col_+2, row_, 'Currency');
   As_Xlsx.CellS (col_+3, row_, 'Amount');
   
   FOR r_ IN get_entities LOOP
      row_ := row_ + 1;
      As_Xlsx.CellS (col_,   row_, r_.identity_type);
      As_Xlsx.CellS (col_+1, row_, r_.identity);
      As_Xlsx.CellS (col_+2, row_, r_.currency);
      As_Xlsx.CellN (col_+3, row_, r_.amount);
   END LOOP;

   data_range_ := as_xlsx.tp_cell_range (
      defined_name => 'SystemData',
      sheet_id     => sheet_,
      tl           => as_xlsx.tp_cell_loc (
         c => col_, r => init_row_, fixc => true, fixr => true
      ),
      br           => as_xlsx.tp_cell_loc (
         c => col_end_, r => row_, fixc => true, fixr => true
      )
   );
   As_Xlsx.Print_Range (data_range_);
   rollup_cols_(1) := 'col';
   rollup_cols_(3) := 'col';
   rollup_cols_(4) := 'sum';

   As_Xlsx.Set_Column_Width (col_,   15, sheet_);
   As_Xlsx.Set_Column_Width (col_+1, 15, sheet_);
   As_Xlsx.Set_Column_Width (col_+2, 15, sheet_);
   As_Xlsx.Set_Column_Width (col_+3, 15, sheet_);
   As_Xlsx.Defined_Name (data_range_);

   As_Xlsx.Add_Pivot_Table (
      cache_id_       => null,
      data_range_     => data_range_,
      rollup_cols_    => rollup_cols_,
      pivot_name_     => 'My Pivot',
      add_to_sheet_   => sheet_,
      new_sheet_name_ => null
   );

   --blob_ := As_Xlsx.Finish;
   file_name_ := file_start_ || 'ImageHypCommNm' || file_end_;
   As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
   Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

END;
0
3
range_.tl.c
range_.br.c
rollup_type_
