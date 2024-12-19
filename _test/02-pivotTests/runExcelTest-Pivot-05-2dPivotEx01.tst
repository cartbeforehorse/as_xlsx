PL/SQL Developer Test script 3.0
93
DECLARE
   xl_blob_     BLOB;
   create_file_ CONSTANT BOOLEAN      := false;
   test_name_   CONSTANT VARCHAR2(30) := 'PivotTable2d';
   file_start_  CONSTANT VARCHAR2(20) := 't05_';
   file_end_    CONSTANT VARCHAR2(20) := to_char(sysdate,'YYYYMMDD-HH24MI');
   file_name_   VARCHAR2(60);
   sheet_       PLS_INTEGER := 1;
   col_         PLS_INTEGER := 1;
   row_         PLS_INTEGER := 2;
   pt_col_      PLS_INTEGER := 2;
   pt_row_      PLS_INTEGER := 18;
   init_row_    PLS_INTEGER := row_;
   data_range_  as_xlsx.tp_cell_range;
   blob_        BLOB;
   cache_id_    PLS_INTEGER;
   loc_         as_xlsx.tp_cell_loc;
   piv_axes_    as_xlsx.tp_pivot_axes := as_xlsx.tp_pivot_axes (
      vrollups    => as_xlsx.tp_pivot_cols(),
      hrollups    => as_xlsx.tp_pivot_cols(),
      filter_cols => as_xlsx.tp_pivot_cols(),
      col_agg_fns => as_xlsx.tp_col_agg_fns()
   );

   CURSOR get_entities IS
      SELECT e.company, e.identity_type, e.identity, e.category, e.currency, e.amount, e.tax
      FROM   entities5dim_tab e;

BEGIN

   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Base Data');

   -- Create data first
   As_Xlsx.CellS (col_,   row_, 'Company');
   As_Xlsx.CellS (col_+1, row_, 'Identity Type');
   As_Xlsx.CellS (col_+2, row_, 'Identity');
   As_Xlsx.CellS (col_+3, row_, 'Category');
   As_Xlsx.CellS (col_+4, row_, 'Currency');
   As_Xlsx.CellS (col_+5, row_, 'Amount');
   As_Xlsx.CellS (col_+6, row_, 'Tax');
   FOR r_ IN get_entities LOOP
      row_ := row_ + 1;
      As_Xlsx.CellS (col_,   row_, r_.company);
      As_Xlsx.CellS (col_+1, row_, r_.identity_type);
      As_Xlsx.CellS (col_+2, row_, r_.identity);
      As_Xlsx.CellS (col_+3, row_, r_.category);
      As_Xlsx.CellS (col_+4, row_, r_.currency);
      As_Xlsx.CellN (col_+5, row_, r_.amount);
      As_Xlsx.CellN (col_+6, row_, r_.tax);
   END LOOP;
   data_range_ := as_xlsx.tp_cell_range (
      defined_name => 'SystemData', -- will create a "defined name" instance, can be commented out
      sheet_id     => sheet_,
      tl           => as_xlsx.tp_cell_loc (col_, init_row_, true, true),
      br           => as_xlsx.tp_cell_loc (col_+6, row_, true, true)
   );
   --As_Xlsx.Defined_Name (data_range_);
   --As_Xlsx.Set_Column_Width (col_,   15, sheet_);
   --As_Xlsx.Set_Column_Width (col_+1, 15, sheet_);
   --As_Xlsx.Set_Column_Width (col_+2, 15, sheet_);
   --As_Xlsx.Set_Column_Width (col_+3, 15, sheet_);

   -------------------------------------
   -- ***
   -- *** HERE ARE THE COLUMN ROLLUPS
   -- ***
   piv_axes_.hrollups       := as_xlsx.tp_pivot_cols(4, 5);
   piv_axes_.vrollups       := as_xlsx.tp_pivot_cols(1, 3);
   piv_axes_.col_agg_fns(1) := as_xlsx.tp_agg_fn (colid => 6, agg_fn => 'sum');
   --piv_axes_.col_agg_fns(2) := as_xlsx.tp_agg_fn (colid => 4, agg_fn => 'count');

   loc_ := as_xlsx.tp_cell_loc (c => pt_col_, r => pt_row_);
   As_Xlsx.Add_Pivot_Table (
      cache_id_       => cache_id_,
      src_data_range_ => data_range_,
      pivot_axes_     => piv_axes_,
      location_tl_    => loc_,
      pivot_name_     => 'AutoPivot01',
      add_to_sheet_   => sheet_
   );
   --As_Xlsx.Set_Column_Width (pt_col_, 15, sheet_);
   --As_Xlsx.Set_Column_Width (pt_col_ + 1, 15, sheet_);

   IF not create_file_ THEN
      xl_blob_ := As_Xlsx.Finish;
   ELSE
      file_name_ := file_start_ || test_name_ || '_' || file_end_ || '.xlsx';
      As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
      Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');
   END IF;

END;
0
14
is_h_leaf_
h_depth_
v_depth_
shared_item_
h_level_
v_level_
col_id_
col_name_
rg_row_start_
rg_row_end_
sum_val_


