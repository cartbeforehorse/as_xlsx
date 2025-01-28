PL/SQL Developer Test script 3.0
102
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '01-basicTests';
   test_name_   CONSTANT VARCHAR2(30) := 'BordersAndFormulas';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(60)          := test_level_ || '-' || test_name_ || file_end_;
   xl_blob_     BLOB;
   sheet_       PLS_INTEGER := 1;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 3;
   init_row_    PLS_INTEGER := row_;
   data_range_  nyce_xlsx.tp_cell_range;
   sum_         NUMBER;
   range_txt_   VARCHAR2(100);
   params_      nyce_xlsx.params_arr := nyce_xlsx.params_arr();

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

BEGIN

   -- Image File
   Nyce_Xlsx.Init_Workbook;

   params_(1) := nyce_xlsx.param_rec (
      param_name      => 'Company Id',
      param_value     => 'HOLDINGS',
      additional_info => 'The company for which this report shows data'
   );
   params_(2) := nyce_xlsx.param_rec (
      param_name      => 'Invoice Id',
      param_value     => '123456789',
      additional_info => 'Invoice held by the customer'
   );
   Nyce_Xlsx.Create_Params_Sheet (
      report_name_ => file_name_,
      params_      => params_,
      show_user_   => true,
      sheet_       => sheet_,
      extra_blurb_ => q'[This output tests the following functionality:
 - Parameters/overview/cover sheet
 - Different color/style of border on each side of cell 
 - Formulas (at least "sum")
 - Merge cells
 - Wrapping of text in a cell, such as seeing what happens when this particular line of text that you are now reading in the moment, wraps over the end of the cell.
 - Border ranges (and verify that they don't overwrite existing content)]'
   );

   sheet_ := Nyce_Xlsx.New_Sheet ('Parameters');
   nyce_xlsx.bdrs_('mixedBdr') := Nyce_Xlsx.Get_Border (
      'dotted', 'thick', 'slantDashDot', 'dashDotDot',
      'FFEE6634', 'FF00FF00', 'FF458609', 'FF2010F0'
   );
   Nyce_Xlsx.CellS (
      2, 2, 'Here''s a list of numbers',
      fontName_ => 'bld_wht', fillName_ => 'dk_green', borderName_ => 'mixedBdr'
   );
   Nyce_Xlsx.CellN (2, 3, 4);
   sum_ := sum_ + 4;
   Nyce_Xlsx.CellN (2, 4, 5);
   sum_ := sum_ + 5;
   Nyce_Xlsx.CellN (2, 5, 7);
   sum_ := sum_ + 7;
   Nyce_Xlsx.CellN (2, 6, 3);
   sum_ := sum_ + 3;
   Nyce_Xlsx.CellN (2, 7, 3);
   sum_ := sum_ + 3;
   Nyce_Xlsx.CellN (2, 8, 8);
   sum_ := sum_ + 8;
   Nyce_Xlsx.CellN (2, 9, 7);
   sum_ := sum_ + 7;

   data_range_.sheet_id     := sheet_;
   data_range_.tl           := Nyce_Xlsx.tp_cell_loc (2, 3, true, true);
   data_range_.br           := Nyce_Xlsx.tp_cell_loc (2, 9, true, true);
   data_range_.defined_name := 'MyDataSource';
   range_txt_ := 'sum(' || Nyce_Xlsx.Alfan_Range (data_range_) || ')';

   Nyce_Xlsx.CellN (2, 10, sum_, range_txt_, fontName_ => 'bld_lg');
   Nyce_Xlsx.Add_Border_To_Range (2, 3, 2, 9, 'thick', 'FF663399');
   Nyce_Xlsx.Add_Border_To_Range (2, 10, 2, 10, 'thick', 'FFFF0000');
   Nyce_Xlsx.CellS (4, 3, 'Cell 1');
   Nyce_Xlsx.CellN (5, 3, 2);
   Nyce_Xlsx.CellN (4, 4, 3);
   Nyce_Xlsx.CellS (5, 4, 'Cell 4');
   Nyce_Xlsx.Add_Border_To_Range (5, 4, 10, 9, 'thick', 'FF00FF00');
   Nyce_Xlsx.Add_Border_To_Range (10, 11, 12, 12, 'thick', 'FF000000');
   Nyce_Xlsx.Add_Border_To_Range (7, 6, 11, 10, 'thick', 'FF0000FF');

   Nyce_Xlsx.Defined_Name (data_range_);

   IF not create_file_ THEN
      xl_blob_ := Nyce_Xlsx.Finish;
      Dbms_Output.Put_Line ('Test script finished: ' || test_level_ || '-' || test_name_);
   ELSE
      Nyce_Xlsx.Save ('EXCEL_OUT', file_name_);
      Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');
   END IF;

END;
1
output_file
1
yes
5
3
range_.tl.c
range_.br.c
rollup_type_
