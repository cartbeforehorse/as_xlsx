PL/SQL Developer Test script 3.0
49
DECLARE
   file_end_   CONSTANT VARCHAR2(20) := Cbh_Utils_API.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_start_ CONSTANT VARCHAR2(20) := 'TestOut_';
   file_name_           VARCHAR2(60);
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
   As_Xlsx.Defined_Name (2, 10, 3, 13, 'CustomerData');

   file_name_ := file_start_ || 'ImageHypCommNm' || file_end_;
   As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
   Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

END Make;
0
2
s_
img_width_
