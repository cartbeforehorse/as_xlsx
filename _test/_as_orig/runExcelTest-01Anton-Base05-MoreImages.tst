PL/SQL Developer Test script 3.0
119
DECLARE
   create_file_ CONSTANT BOOLEAN      := lower(nvl(:output_file,'no')) = 'yes';
   test_level_  CONSTANT VARCHAR2(30) := '01-basicTests';
   test_name_   CONSTANT VARCHAR2(30) := 'AS05-MoreImages';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(60);
   xl_blob_     BLOB;
   sheet_       PLS_INTEGER;
   col_         PLS_INTEGER := 2;
   col_end_     PLS_INTEGER := col_ + 3;
   row_         PLS_INTEGER := 3;
   init_row_    PLS_INTEGER := row_;

   CURSOR get_entities IS
      SELECT e.identity_type, e.identity, e.currency, e.amount
      FROM   entities_tab e;

   FUNCTION Load_File (
      filename_ IN VARCHAR2 ) RETURN BLOB
   IS
      dir_      VARCHAR2(100) := 'EXCEL_STORAGE';
      img_blob_ BLOB          := empty_blob();
      bfile_    BFILE         := bFileName (dir_, filename_);
   BEGIN
      Dbms_Lob.fileOpen (bfile_);
      Dbms_Lob.createTemporary (img_blob_, true);
      Dbms_Lob.loadFromFile (img_blob_, bfile_, Dbms_Lob.getLength(bfile_));
      Dbms_Lob.fileClose (bfile_);
      RETURN img_blob_;
   END Load_File;

BEGIN

   -- Image File
   As_Xlsx.New_Sheet ('Parameters');
   As_Xlsx.Cell (2, 2, 'dummy data');
   As_Xlsx.Hyperlink (2, 3, 'https://cartbeforehorse.com', 'click me');
   As_Xlsx.Comment (2, 2, 'This is a silly dingle dongle', 'Bob the Builder', 300, 200, 1);
   As_Xlsx.Add_Image (
      p_col => 4, p_row => 2, p_img => Load_File('signature.jpg'),
      p_name        => 'Sig Name',
      p_title       => 'Sig Title',
      p_description => 'Sig Description',
      p_scale => 0.1, p_sheet => 1
   );

   As_Xlsx.Cell (2, 10, 'Customer Id');
   As_Xlsx.Cell (3, 10, 'Customer Name');
   As_Xlsx.Cell (2, 11, '100103');
   As_Xlsx.Cell (3, 11, 'Charlie the grey squirel');
   As_Xlsx.Cell (2, 12, '100103');
   As_Xlsx.Cell (3, 12, 'Casablanka (the city)');
   As_Xlsx.Cell (2, 13, '100103');
   As_Xlsx.Cell (3, 13, 'Bing Bong the bouncing compnay');
   As_Xlsx.Defined_Name (
      p_tl_col => 2, p_tl_row => 10, p_br_col => 3, p_br_row => 13,
      p_name => 'CustomerData', p_sheet => 1
   );

   As_Xlsx.New_Sheet ('Number Two');
   As_Xlsx.Add_Image (
      p_col => 8, p_row => 2, p_img => Load_File('excel.png'),
      p_name        => 'Excel Image Name',
      p_title       => 'Excel Logo Title',
      p_description => 'Excel Logo Description',
      p_scale => 0.1, p_sheet => 2
   );
   As_Xlsx.Add_Image (
      p_col => 2, p_row => 10, p_img => Load_File('bitmap-green.bmp'),
      p_name        => 'Bitmap Image Name',
      p_title       => 'Bitmap Title',
      p_description => 'Bitmap Description',
      p_scale => 0.5, p_sheet => 2
   );

   As_Xlsx.New_Sheet ('Data');

   As_Xlsx.Cell (col_,   row_, 'Identity Type', p_sheet => 3);
   As_Xlsx.Cell (col_+1, row_, 'Identity', p_sheet => 3);
   As_Xlsx.Cell (col_+2, row_, 'Currency', p_sheet => 3);
   As_Xlsx.Cell (col_+3, row_, 'Amount', p_sheet => 3);
   
   FOR r_ IN get_entities LOOP
      row_ := row_ + 1;
      As_Xlsx.Cell (col_,   row_, r_.identity_type, p_sheet => 3);
      As_Xlsx.Cell (col_+1, row_, r_.identity, p_sheet => 3);
      As_Xlsx.Cell (col_+2, row_, r_.currency, p_sheet => 3);
      As_Xlsx.Cell (col_+3, row_, r_.amount, p_sheet => 3);
   END LOOP;

   As_Xlsx.Set_Column_Width (col_,   15, 3);
   As_Xlsx.Set_Column_Width (col_+1, 15, 3);
   As_Xlsx.Set_Column_Width (col_+2, 15, 3);
   As_Xlsx.Set_Column_Width (col_+3, 15, 3);
   As_Xlsx.Defined_Name (
      p_tl_col => col_, p_tl_row => init_row_, p_br_col => col_+3, p_br_row => row_,
      p_name => 'MyDataSource', p_sheet => 3
   );


   As_Xlsx.New_Sheet ('Number Four');
   As_Xlsx.Add_Image (
      p_col => 2, p_row => 2, p_img => Load_File('excel.png'),
      p_name        => 'Excel Image Name 2',
      p_title       => 'Excel Title 2',
      p_description => 'Excel Description 2',
      p_scale => 0.1, p_sheet => 4
   );

   IF not create_file_ THEN
      xl_blob_ := As_Xlsx.Finish;
      Dbms_Output.Put_Line ('Test script finished: ' || test_level_ || '-' || test_name_);
   ELSE
      file_name_ := test_level_ || '-' || test_name_ || file_end_;
      As_Xlsx.Save ('EXCEL_OUT', file_name_);
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
