select xl.sheet_nr, xl.sheet_name, xl.row_nr, xl.col_nr, xl.cell, xl.cell_type, xl.string_val, xl.number_val, xl.date_val, xl.formula, xl.string_len
from   table(Nyce_Xlsx.Read_Xl(dir_ => 'EXCEL_OUT', filename_ => '02-tables-02-MultiTablesOnPage_20250203-2017.xlsx')) xl;

select xl.sheet_nr, xl.sheet_name, xl.row_nr, xl.col_nr, xl.cell, xl.cell_type, xl.string_val, xl.number_val, xl.date_val, xl.formula, xl.string_len
from   table(Nyce_Xlsx.Read_Xl(dir_ => 'EXCEL_OUT', filename_ => '01-basicTests-NumFormats_20250203-2045.xlsx')) xl;

  --FUNCTION Read_Xl (
  -- excel_          IN BLOB     := null,
  -- sheets_         IN VARCHAR2 := null,
  -- cell_           IN VARCHAR2 := null,
  -- include_clobs_  IN VARCHAR2 := null,
  -- add_empty_cols_ IN VARCHAR2 := null,
  -- dir_            IN VARCHAR2 := null,
  -- filename_       IN VARCHAR2 := null ) RETURN tp_all_cells PIPELINED;
