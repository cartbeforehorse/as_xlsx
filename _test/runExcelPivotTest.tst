PL/SQL Developer Test script 3.0
79
DECLARE
   file_end_   CONSTANT VARCHAR2(20) := Cbh_Utils_API.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_start_ CONSTANT VARCHAR2(20) := 'TestOut_';
   file_name_           VARCHAR2(60);
BEGIN

   -- Image File
   As_Xlsx.Init_Workbook;
   As_Xlsx.Set_Sheet_Name (1, 'Data');

   -- headings
   As_Xlsx.CellS (2, 2, 'Identity Type');
   As_Xlsx.CellS (3, 2, 'Identity');
   As_Xlsx.CellS (4, 2, 'Currency');
   As_Xlsx.CellS (5, 2, 'Amount');
   -- row 1
   As_Xlsx.CellS (2, 3, 'Supplier');
   As_Xlsx.CellS (3, 3, 'Supp Dude');
   As_Xlsx.CellS (4, 3, 'EUR');
   As_Xlsx.CellN (5, 3, 1012.69);
   -- row 2
   As_Xlsx.CellS (2, 4, 'Supplier');
   As_Xlsx.CellS (3, 4, 'George Doors');
   As_Xlsx.CellS (4, 4, 'GBP');
   As_Xlsx.CellN (5, 4, 1200);
   -- row 3
   As_Xlsx.CellS (2, 5, 'Customer');
   As_Xlsx.CellS (3, 5, 'Car Bits');
   As_Xlsx.CellS (4, 5, 'GBP');
   As_Xlsx.CellN (5, 5, 600.78);
   -- row 4; Supplier  Green Engines  GBP 354
   As_Xlsx.CellS (2, 6, 'Supplier'); --
   As_Xlsx.CellS (3, 6, 'Green Engines');
   As_Xlsx.CellS (4, 6, 'GBP');
   As_Xlsx.CellN (5, 6, 345);
   -- row 5: Customer  Ford  SEK 789.54
   As_Xlsx.CellS (2, 7, 'Customer'); --
   As_Xlsx.CellS (3, 7, 'Ford');
   As_Xlsx.CellS (4, 7, 'SEK');
   As_Xlsx.CellN (5, 7, 789.54);
   -- row 6: Customer  Ford  GBP 416.8
   As_Xlsx.CellS (2, 8, 'Customer'); --
   As_Xlsx.CellS (3, 8, 'Ford');
   As_Xlsx.CellS (4, 8, 'GBP');
   As_Xlsx.CellN (5, 8, 416.8);
   -- row 7: Customer  Volvo   GBP   168.56
   As_Xlsx.CellS (2, 9, 'Customer'); --
   As_Xlsx.CellS (3, 9, 'Volvo');
   As_Xlsx.CellS (4, 9, 'GBP');
   As_Xlsx.CellN (5, 9, 168.56);
   -- row 8: Customer  Ford  USD 56.4
   As_Xlsx.CellS (2, 10, 'Customer'); --
   As_Xlsx.CellS (3, 10, 'Ford');
   As_Xlsx.CellS (4, 10, 'USD');
   As_Xlsx.CellN (5, 10, 56.4);
   -- row 9: Customer  Car Bits   EUR   333.34
   As_Xlsx.CellS (2, 11, 'Customer'); --
   As_Xlsx.CellS (3, 11, 'Car Bits');
   As_Xlsx.CellS (4, 11, 'EUR');
   As_Xlsx.CellN (5, 11, 333.34);
   -- row 10: Supplier  George Doors EUR  6788
   As_Xlsx.CellS (2, 12, 'Supplier'); --
   As_Xlsx.CellS (3, 12, 'George Doors');
   As_Xlsx.CellS (4, 12, 'EUR');
   As_Xlsx.CellN (5, 12, 6788);

   As_Xlsx.Defined_Name (2, 2, 5, 12, 'SourceData');
   As_Xlsx.Set_Column_Width (2, 12);
   As_Xlsx.Set_Column_Width (3, 20);

   -- Sheet 2
   As_Xlsx.New_Sheet ('Pivoted Data');
   As_Xlsx.Add_Pivot ('My Pivot', 2);

   file_name_ := file_start_ || 'Pivot' || file_end_;
   As_Xlsx.Save (As_Xlsx.Finish, 'EXCEL_OUT', file_name_);
   Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');

END Make;
0
2
s_
img_width_
