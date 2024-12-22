PL/SQL Developer Test script 3.0
89
DECLARE
   test_name_   CONSTANT VARCHAR2(30) := 'PivotTableCol';
   file_start_  CONSTANT VARCHAR2(20) := 'TestOut_';
   file_end_    CONSTANT VARCHAR2(20) := to_char(sysdate,'YYYYMMDD-HH24MI');
   file_name_   VARCHAR2(60);
   sheet_       PLS_INTEGER := 1;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 2;
   init_row_    PLS_INTEGER := row_;
   data_range_  nyce_xlsx.tp_cell_range;
   blob_        BLOB;
   cache_id_    PLS_INTEGER;
   loc_         nyce_xlsx.tp_cell_loc := nyce_xlsx.tp_cell_loc (c => 8, r => 2);
   piv_axes_    nyce_xlsx.tp_pivot_axes := nyce_xlsx.tp_pivot_axes (
      vrollups    => nyce_xlsx.tp_pivot_cols(),
      hrollups    => nyce_xlsx.tp_pivot_cols(),
      filter_cols => nyce_xlsx.tp_pivot_cols(),
      col_agg_fns => nyce_xlsx.tp_col_agg_fns()
   );
   arr_         Nyce_Xlsx.tp_pivot_cols;

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

BEGIN

   FOR i_ IN 1 .. 3 LOOP

      --Nyce_Xlsx.Clear_Workbook;
      cache_id_ := null;
      row_      := 2;

      arr_(i_) := i_;
      dbms_output.put_line ('loop: ' || i_);
      piv_axes_.vrollups       := arr_; --Nyce_Xlsx.tp_pivot_cols(1, 3);
      piv_axes_.col_agg_fns(4) := 'sum';

      -- Image File
      Nyce_Xlsx.Init_Workbook;
      Nyce_Xlsx.Set_Sheet_Name (1, 'Base Data');

      Nyce_Xlsx.CellS (col_,   row_, 'Identity Type');
      Nyce_Xlsx.CellS (col_+1, row_, 'Identity');
      Nyce_Xlsx.CellS (col_+2, row_, 'Currency');
      Nyce_Xlsx.CellS (col_+3, row_, 'Amount');

      FOR r_ IN get_entities LOOP
         row_ := row_ + 1;
         Nyce_Xlsx.CellS (col_,   row_, r_.identity_type);
         Nyce_Xlsx.CellS (col_+1, row_, r_.identity);
         Nyce_Xlsx.CellS (col_+2, row_, r_.currency);
         Nyce_Xlsx.CellN (col_+3, row_, r_.amount);
      END LOOP;

      data_range_ := nyce_xlsx.tp_cell_range (
         defined_name => 'SystemData', -- will create a "defined name" instance, can be commented out
         sheet_id     => sheet_,
         tl           => nyce_xlsx.tp_cell_loc (col_, init_row_, true, true),
         br           => nyce_xlsx.tp_cell_loc (col_end_, row_, true, true)
      );
      --Nyce_Xlsx.Defined_Name (data_range_);
      --Nyce_Xlsx.Print_Range (data_range_); -- debug

      Nyce_Xlsx.Set_Column_Width (col_,   15, sheet_);
      Nyce_Xlsx.Set_Column_Width (col_+1, 15, sheet_);
      Nyce_Xlsx.Set_Column_Width (col_+2, 15, sheet_);
      Nyce_Xlsx.Set_Column_Width (col_+3, 15, sheet_);

      Nyce_Xlsx.Add_Pivot_Table (
         cache_id_       => cache_id_,
         src_data_range_ => data_range_,
         pivot_axes_     => piv_axes_,
         location_tl_    => loc_,
         pivot_name_     => 'AutoPivot',
         add_to_sheet_   => sheet_
      );
      Nyce_Xlsx.Set_Column_Width (8, 15, sheet_);
      Nyce_Xlsx.Set_Column_Width (9, 15, sheet_);

      --blob_ := Nyce_Xlsx.Finish;
      file_name_ := file_start_ || test_name_ || '_' || file_end_ || '-' || i_ || '.xlsx';
      Nyce_Xlsx.Save (Nyce_Xlsx.Finish, 'EXCEL_OUT', file_name_);
      Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

   END LOOP;

END;
0
14
value_












