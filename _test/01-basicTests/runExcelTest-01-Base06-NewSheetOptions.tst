PL/SQL Developer Test script 3.0
120
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '01-basicTests';
   test_name_   CONSTANT VARCHAR2(30) := 'NewSheetOptions';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(60);
   xl_blob_     BLOB;
   sheet_       PLS_INTEGER;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 3;
   init_row_    PLS_INTEGER := row_;
   data_range_  nyce_xlsx.tp_cell_range;

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

BEGIN

   -- Image File
   Nyce_Xlsx.Init_Workbook;
   Nyce_Xlsx.Set_Sheet_Name (1, 'Parameters');
   Nyce_Xlsx.CellS (2, 2, 'dummy data');
   Nyce_Xlsx.Hyperlink (2, 3, 'https://cartbeforehorse.com', 'click me');
   Nyce_Xlsx.Comment (2, 2, 'This is a silly dingle dongle', 'Bob the Builder', 300, 200, 1);
   Nyce_Xlsx.Load_Image (
      col_         => 4,
      row_         => 2,
      dir_         => 'EXCEL_STORAGE',
      filename_    => 'signature.jpg',
      name_        => 'Excel Image Name',
      title_       => 'Excel Logo Title',
      description_ => 'Excel Logo Description',
      scale_       => 0.1,
      sheet_       => 1
   );
   Nyce_Xlsx.CellS (2, 10, 'Customer Id');
   Nyce_Xlsx.CellS (3, 10, 'Customer Name');
   Nyce_Xlsx.CellS (2, 11, '100103');
   Nyce_Xlsx.CellS (3, 11, 'Charlie the grey squirel');
   Nyce_Xlsx.CellS (2, 12, '100103');
   Nyce_Xlsx.CellS (3, 12, 'Casablanka (the city)');
   Nyce_Xlsx.CellS (2, 13, '100103');
   Nyce_Xlsx.CellS (3, 13, 'Bing Bong the bouncing compnay');
   Nyce_Xlsx.Defined_Name ('CustomerData', 2, 10, 3, 13, sheet_ => 1);

   sheet_ := Nyce_Xlsx.New_Sheet ('Number Two', show_headers_ => false);
   Nyce_Xlsx.Load_Image (
      col_         => 8,
      row_         => 2,
      dir_         => 'EXCEL_STORAGE',
      filename_    => 'excel.png',
      name_        => 'Excel Image Name',
      title_       => 'Excel Logo Title',
      description_ => 'Excel Logo Description',
      scale_       => 0.1,
      sheet_       => sheet_
   );
   Nyce_Xlsx.Load_Image (
      col_         => 2,
      row_         => 10,
      dir_         => 'EXCEL_STORAGE',
      filename_    => 'bitmap-green.bmp',
      name_        => 'Bitmap Image Name',
      title_       => 'Bitmap Title',
      description_ => 'A splash for bitmaps',
      scale_       => 0.5,
      sheet_       => sheet_
   );

   sheet_ := Nyce_Xlsx.New_Sheet ('Data', grid_colour_ix_ => 6);

   Nyce_Xlsx.CellS (col_,   row_, 'Identity Type', sheet_ => sheet_);
   Nyce_Xlsx.CellS (col_+1, row_, 'Identity', sheet_ => sheet_);
   Nyce_Xlsx.CellS (col_+2, row_, 'Currency', sheet_ => sheet_);
   Nyce_Xlsx.CellS (col_+3, row_, 'Amount', sheet_ => sheet_);
   
   FOR r_ IN get_entities LOOP
      row_ := row_ + 1;
      Nyce_Xlsx.CellS (col_,   row_, r_.identity_type, sheet_ => sheet_);
      Nyce_Xlsx.CellS (col_+1, row_, r_.identity, sheet_ => sheet_);
      Nyce_Xlsx.CellS (col_+2, row_, r_.currency, sheet_ => sheet_);
      Nyce_Xlsx.CellN (col_+3, row_, r_.amount, sheet_ => sheet_);
   END LOOP;

   Nyce_Xlsx.Set_Column_Width (col_,   15, sheet_);
   Nyce_Xlsx.Set_Column_Width (col_+1, 15, sheet_);
   Nyce_Xlsx.Set_Column_Width (col_+2, 15, sheet_);
   Nyce_Xlsx.Set_Column_Width (col_+3, 15, sheet_);

   data_range_.sheet_id     := sheet_;
   data_range_.tl           := Nyce_Xlsx.tp_cell_loc (col_, init_row_, true, true);
   data_range_.br           := Nyce_Xlsx.tp_cell_loc (col_ + 3, row_, true, true);
   data_range_.defined_name := 'MyDataSource';
   Nyce_Xlsx.Defined_Name (data_range_);

   sheet_ := Nyce_Xlsx.New_Sheet ('Number Four', 'FFFFFF00', false, grid_colour_ix_ => 4);
   Nyce_Xlsx.Load_Image (
      col_         => 2,
      row_         => 2,
      dir_         => 'EXCEL_STORAGE',
      filename_    => 'excel.png',
      name_        => 'Excel Image Name 2',
      title_       => 'Excel Logo Title 2',
      description_ => 'Excel Logo Duplicate',
      scale_       => 0.1,
      sheet_       => sheet_
   );

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
0
5
3
range_.tl.c
range_.br.c
rollup_type_
