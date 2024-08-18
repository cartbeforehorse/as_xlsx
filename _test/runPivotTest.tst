PL/SQL Developer Test script 3.0
75
DECLARE
   file_end_    CONSTANT VARCHAR2(20) := Cbh_Utils_API.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_start_  CONSTANT VARCHAR2(20) := 'TestOut_';
   file_name_   VARCHAR2(60);
   sheet_       PLS_INTEGER := 1;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 2;
   init_row_    PLS_INTEGER := row_;
   data_range_  as_xlsx.tp_cell_range;
   blob_        BLOB;
   cache_id_    PLS_INTEGER;
   loc_         as_xlsx.tp_cell_loc := as_xlsx.tp_cell_loc (c => 8, r => 2);
   piv_axes_    as_xlsx.tp_pivot_axes := as_xlsx.tp_pivot_axes (
      vrollups    => as_xlsx.tp_pivot_cols(),
      hrollups    => as_xlsx.tp_pivot_cols(),
      filter_cols => as_xlsx.tp_pivot_cols(),
      col_agg_fns => as_xlsx.tp_col_agg_fns()
   );

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

BEGIN

   piv_axes_.vrollups       := as_xlsx.tp_pivot_cols(1, 3);
   piv_axes_.col_agg_fns(4) := 'sum';

   -- Image File
   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Base Data');

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
      --defined_name => 'SystemData', -- will create a "defined name" instance
      sheet_id     => sheet_,
      tl           => as_xlsx.tp_cell_loc (col_, init_row_, true, true),
      br           => as_xlsx.tp_cell_loc (col_end_, row_, true, true)
   );
   --As_Xlsx.Defined_Name (data_range_);
   As_Xlsx.Print_Range (data_range_); -- debug

   As_Xlsx.Set_Column_Width (col_,   15, sheet_);
   As_Xlsx.Set_Column_Width (col_+1, 15, sheet_);
   As_Xlsx.Set_Column_Width (col_+2, 15, sheet_);
   As_Xlsx.Set_Column_Width (col_+3, 15, sheet_);

   As_Xlsx.Add_Pivot_Table (
      cache_id_       => cache_id_,
      src_data_range_ => data_range_,
      pivot_axes_     => piv_axes_,
      location_tl_    => loc_,
      pivot_name_     => 'OsianPivot',
      add_to_sheet_   => sheet_
   );

   --blob_ := As_Xlsx.Finish;
   file_name_ := file_start_ || 'PivotTest' || file_end_;
   As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
   Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

END;
0
14
value_












