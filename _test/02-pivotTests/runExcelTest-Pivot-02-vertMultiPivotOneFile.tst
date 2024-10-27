PL/SQL Developer Test script 3.0
93
DECLARE
   test_name_   CONSTANT VARCHAR2(30) := 'PivotTableVerticalMulticache';
   file_start_  CONSTANT VARCHAR2(20) := 't02_';
   file_end_    CONSTANT VARCHAR2(20) := to_char(sysdate,'YYYYMMDD-HH24MI');
   file_name_   VARCHAR2(60);
   sheet_       PLS_INTEGER := 1;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 2;
   pt_col_      PLS_INTEGER;
   pt_row_      PLS_INTEGER;
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
   arr_         as_xlsx.tp_pivot_cols;

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

BEGIN

   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Base Data');

   -- Create data first
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
      defined_name => 'SystemData', -- will create a "defined name" instance, can be commented out
      sheet_id     => sheet_,
      tl           => as_xlsx.tp_cell_loc (col_, init_row_, true, true),
      br           => as_xlsx.tp_cell_loc (col_end_, row_, true, true)
   );
   As_Xlsx.Set_Column_Width (col_,   15, sheet_);
   As_Xlsx.Set_Column_Width (col_+1, 15, sheet_);
   As_Xlsx.Set_Column_Width (col_+2, 15, sheet_);
   As_Xlsx.Set_Column_Width (col_+3, 15, sheet_);
   
   FOR i_ IN 1 .. 3 LOOP

      cache_id_ := null;
      IF i_ = 1 THEN
         sheet_  := i_;
         pt_col_ := 8;
         pt_row_ := 2;
      ELSE
         sheet_  := null;
         pt_col_ := 2;
         pt_row_ := 2;
      END IF;

      arr_(i_) := i_;
      dbms_output.put_line ('loop: ' || i_);
      piv_axes_.vrollups       := arr_; --as_xlsx.tp_pivot_cols(1, 3);
      piv_axes_.col_agg_fns(4) := 'sum';

      loc_ := as_xlsx.tp_cell_loc (c => pt_col_, r => pt_row_);
      As_Xlsx.Add_Pivot_Table (
         cache_id_       => cache_id_,
         src_data_range_ => data_range_,
         pivot_axes_     => piv_axes_,
         location_tl_    => loc_,
         pivot_name_     => 'AutoPivot' || to_char(i_),
         add_to_sheet_   => sheet_
      );
      As_Xlsx.Set_Column_Width (pt_col_, 15, sheet_);
      As_Xlsx.Set_Column_Width (pt_col_ + 1, 15, sheet_);

   END LOOP;

   file_name_ := file_start_ || test_name_ || '_' || file_end_ || '.xlsx';
   As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
   Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

END;
0
14
value_












