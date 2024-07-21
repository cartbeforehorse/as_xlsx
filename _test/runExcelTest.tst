PL/SQL Developer Test script 3.0
111
DECLARE
   file_end_    CONSTANT VARCHAR2(20) := Cbh_Utils_API.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_start_  CONSTANT VARCHAR2(20) := 'TestOut_';
   file_name_   VARCHAR2(60);
   sheet_       PLS_INTEGER;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 3;
   init_row_    PLS_INTEGER := row_;
   data_range_  as_xlsx.tp_cell_range;
   rollup_cols_ as_xlsx.rollup_columns;
   excel_       BLOB;
   gen_file_    BOOLEAN := false;

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

BEGIN

   -- Image File
   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Parameters');
   As_Xlsx.CellS (2, 2, 'dummy data');
   As_Xlsx.Hyperlink (2, 3, 'https://cartbeforehorse.com', 'click me');
   As_Xlsx.Comment (2, 2, 'This is a silly dingle dongle', 'Bob the Builder', 300, 200, 1);
   As_Xlsx.Load_Image (
      col_         => 4,
      row_         => 2,
      dir_         => 'EXCEL_OUT',
      filename_    => 'signature.jpg',
      name_        => 'Excel Image Name',
      title_       => 'Excel Logo Title',
      description_ => 'Excel Logo Description',
      scale_       => 0.1,
      sheet_       => 1
   );
   As_Xlsx.Load_Image (
      col_         => 8,
      row_         => 2,
      dir_         => 'EXCEL_OUT',
      filename_    => 'excel.png',
      name_        => 'Excel Image Name',
      title_       => 'Excel Logo Title',
      description_ => 'Excel Logo Description',
      scale_       => 0.1,
      sheet_       => 1
   );
   As_Xlsx.CellS (2, 10, 'Customer Id');
   As_Xlsx.CellS (3, 10, 'Customer Name');
   As_Xlsx.CellS (2, 11, '100103');
   As_Xlsx.CellS (3, 11, 'Charlie the grey squirel');
   As_Xlsx.CellS (2, 12, '100103');
   As_Xlsx.CellS (3, 12, 'Casablanka (the city_');
   As_Xlsx.CellS (2, 13, '100103');
   As_Xlsx.CellS (3, 13, 'Bing Bong the bouncing compnay');
   As_Xlsx.Defined_Name ('CustomerData', 2, 10, 3, 13, sheet_ => 1);

   sheet_ := As_Xlsx.New_Sheet ('Data and Pivot');

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

   excel_     := As_Xlsx.Finish;
   file_name_ := file_start_ || 'ImageHypCommNm' || file_end_;
   IF gen_file_ THEN
      As_Xlsx.Save (excel_, 'EXCEL_OUT', file_name_);
      Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');
   END IF;

END;
0
3
range_.tl.c
range_.br.c
rollup_type_
