PL/SQL Developer Test script 3.0
90
DECLARE
   create_file_ CONSTANT BOOLEAN      := true;
   xl_blob_     BLOB;
   test_name_   CONSTANT VARCHAR2(30) := 'Tables03-TablesWithOtherRels';
   file_start_  CONSTANT VARCHAR2(20) := 'TestOut_';
   file_end_    CONSTANT VARCHAR2(20) := Nyce_Utils.Rep ('_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   file_name_   VARCHAR2(60);

   col_count_   PLS_INTEGER;
   row_count_   PLS_INTEGER;
   sheet_       PLS_INTEGER;

   PROCEDURE Create_Data (
      col_start_ IN PLS_INTEGER,
      row_start_ IN PLS_INTEGER,
      sh_        IN PLS_INTEGER )
   IS
      col_  PLS_INTEGER := col_start_;
      row_  PLS_INTEGER := row_start_;
      CURSOR get_entities IS
         SELECT e.company, e.identity_type, e.identity, e.category, e.currency, e.amount, e.tax
         FROM   entities5dim_tab e;
   BEGIN
      Nyce_Xlsx.CellS (col_,   row_, 'Company', sheet_ => sh_);
      Nyce_Xlsx.CellS (col_+1, row_, 'Identity Type', sheet_ => sh_);
      Nyce_Xlsx.CellS (col_+2, row_, 'Identity', sheet_ => sh_);
      Nyce_Xlsx.CellS (col_+3, row_, 'Category', sheet_ => sh_);
      Nyce_Xlsx.CellS (col_+4, row_, 'Currency', sheet_ => sh_);
      Nyce_Xlsx.CellS (col_+5, row_, 'Amount', sheet_ => sh_);
      Nyce_Xlsx.CellS (col_+6, row_, 'Tax', sheet_ => sh_);
      FOR r_ IN get_entities LOOP
         row_ := row_ + 1;
         Nyce_Xlsx.CellS (col_,   row_, r_.company, sheet_ => sh_);
         Nyce_Xlsx.CellS (col_+1, row_, r_.identity_type, sheet_ => sh_);
         Nyce_Xlsx.CellS (col_+2, row_, r_.identity, sheet_ => sh_);
         Nyce_Xlsx.CellS (col_+3, row_, r_.category, sheet_ => sh_);
         Nyce_Xlsx.CellS (col_+4, row_, r_.currency, sheet_ => sh_);
         Nyce_Xlsx.CellN (col_+5, row_, r_.amount, sheet_ => sh_);
         Nyce_Xlsx.CellN (col_+6, row_, r_.tax, sheet_ => sh_);
      END LOOP;
   END Create_Data;

BEGIN

   sheet_ := Nyce_Xlsx.New_Sheet ('Data and Tables');
   Create_Data (2, 2, sheet_);
   Create_Data (4, 18, sheet_);

   -- TableStyleLight1;TableStyleLight21;TableStyleMedium1;TableStyleMedium28;TableStyleDark1;TableStyleDark11
   Nyce_Xlsx.Set_Table (
      tbl_range_ => nyce_xlsx.tp_cell_range (
         sheet_id => sheet_,
         tl       => nyce_xlsx.tp_cell_loc (2, 2),
         br       => nyce_xlsx.tp_cell_loc (8, 16)
      ),
      style_     => 'TableStyleLight21',
      tbl_name_  => 'Table01'
   );
   Nyce_Xlsx.Set_Table (
      tbl_range_ => nyce_xlsx.tp_cell_range (
         sheet_id => sheet_,
         tl       => nyce_xlsx.tp_cell_loc (4, 18),
         br       => nyce_xlsx.tp_cell_loc (10, 32)
      ),
      style_     => 'TableStyleDark1',
      tbl_name_  => 'Table02'
   );
   Nyce_Xlsx.Hyperlink (10, 2, 'http://nyce.software', 'NYCE Link', sheet_);
   Nyce_Xlsx.Comment (10, 2, 'Comment on the hyperlink', 'Osian', 300, 200, 1);
   Nyce_Xlsx.Load_Image (
      col_         => 10,
      row_         => 6,
      dir_         => 'EXCEL_STORAGE',
      filename_    => 'signature.jpg',
      name_        => 'Excel Image Name',
      title_       => 'Excel Logo Title',
      description_ => 'Excel Logo Description',
      scale_       => 0.1,
      sheet_       => 1
   );

   IF not create_file_ THEN
      xl_blob_ := Nyce_Xlsx.Finish;
   ELSE
      file_name_ := file_start_ || test_name_ || file_end_;
      Nyce_Xlsx.Save ('EXCEL_OUT', file_name_);
      Dbms_Output.Put_Line (file_name_ || ' saved to filesystem');
   END IF;

END;
0
3
range_.tl.c
range_.br.c
rollup_type_
