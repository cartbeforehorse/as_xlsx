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
   data_range_  nyce_xlsx.tp_cell_range;
   blob_        BLOB;
   cache_id_    PLS_INTEGER;
   loc_         nyce_xlsx.tp_cell_loc;
   piv_axes_    nyce_xlsx.tp_pivot_axes := nyce_xlsx.tp_pivot_axes (
      vrollups    => nyce_xlsx.tp_pivot_cols(),
      hrollups    => nyce_xlsx.tp_pivot_cols(),
      filter_cols => nyce_xlsx.tp_pivot_cols(),
      col_agg_fns => nyce_xlsx.tp_col_agg_fns()
   );

   CURSOR get_entities IS
      SELECT e.company, e.identity_type, e.identity, e.category, e.currency, e.amount, e.tax
      FROM   entities5dim_tab e;

BEGIN

   Nyce_Xlsx.Init_Workbook;
   Nyce_Xlsx.Set_Sheet_Name (1, 'Base Data');

   -- Create data first
   Nyce_Xlsx.CellS (col_,   row_, 'Company');
   Nyce_Xlsx.CellS (col_+1, row_, 'Identity Type');
   Nyce_Xlsx.CellS (col_+2, row_, 'Identity');
   Nyce_Xlsx.CellS (col_+3, row_, 'Category');
   Nyce_Xlsx.CellS (col_+4, row_, 'Currency');
   Nyce_Xlsx.CellS (col_+5, row_, 'Amount');
   Nyce_Xlsx.CellS (col_+6, row_, 'Tax');
   FOR r_ IN get_entities LOOP
      row_ := row_ + 1;
      Nyce_Xlsx.CellS (col_,   row_, r_.company);
      Nyce_Xlsx.CellS (col_+1, row_, r_.identity_type);
      Nyce_Xlsx.CellS (col_+2, row_, r_.identity);
      Nyce_Xlsx.CellS (col_+3, row_, r_.category);
      Nyce_Xlsx.CellS (col_+4, row_, r_.currency);
      Nyce_Xlsx.CellN (col_+5, row_, r_.amount);
      Nyce_Xlsx.CellN (col_+6, row_, r_.tax);
   END LOOP;
   data_range_ := nyce_xlsx.tp_cell_range (
      defined_name => 'SystemData', -- will create a "defined name" instance, can be commented out
      sheet_id     => sheet_,
      tl           => nyce_xlsx.tp_cell_loc (col_, init_row_, true, true),
      br           => nyce_xlsx.tp_cell_loc (col_+6, row_, true, true)
   );
   --Nyce_Xlsx.Defined_Name (data_range_);
   --Nyce_Xlsx.Set_Column_Width (col_,   15, sheet_);
   --Nyce_Xlsx.Set_Column_Width (col_+1, 15, sheet_);
   --Nyce_Xlsx.Set_Column_Width (col_+2, 15, sheet_);
   --Nyce_Xlsx.Set_Column_Width (col_+3, 15, sheet_);

   -------------------------------------
   -- ***
   -- *** HERE ARE THE COLUMN ROLLUPS
   -- ***
   piv_axes_.hrollups       := nyce_xlsx.tp_pivot_cols(4, 5);
   piv_axes_.vrollups       := nyce_xlsx.tp_pivot_cols(1, 3);
   piv_axes_.col_agg_fns(1) := nyce_xlsx.tp_agg_fn (colid => 6, agg_fn => 'sum');
   --piv_axes_.col_agg_fns(2) := nyce_xlsx.tp_agg_fn (colid => 4, agg_fn => 'count');

   loc_ := nyce_xlsx.tp_cell_loc (c => pt_col_, r => pt_row_);
   Nyce_Xlsx.Add_Pivot_Table (
      cache_id_       => cache_id_,
      src_data_range_ => data_range_,
      pivot_axes_     => piv_axes_,
      location_tl_    => loc_,
      pivot_name_     => 'AutoPivot01',
      add_to_sheet_   => sheet_
   );
   --Nyce_Xlsx.Set_Column_Width (pt_col_, 15, sheet_);
   --Nyce_Xlsx.Set_Column_Width (pt_col_ + 1, 15, sheet_);

   IF not create_file_ THEN
      xl_blob_ := Nyce_Xlsx.Finish;
   ELSE
      file_name_ := file_start_ || test_name_ || '_' || file_end_ || '.xlsx';
      Nyce_Xlsx.Save (Nyce_Xlsx.Finish, 'EXCEL_OUT', file_name_);
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


