PL/SQL Developer Test script 3.0
36
declare
  l_cnt pls_integer;
  l_query sys_refcursor;
begin
  open l_query for
    select date '1900-02-26' + level "Secret Date"
         , to_char( date '1900-02-26' + level, 'yyyy mon dd' ) "Secret String"
    from dual
    connect by level < 8;
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  l_cnt := as_xlsx.query2sheet
             ( p_rc         => l_query
             , p_sheet      => 1
             , p_col        => 5
             , p_row        => 3
             , p_autofilter => true
             , p_date_format => 'yyyy-mmm-dd'
             , p_title      => 'My Secrets'
             , p_title_xfid => as_xlsx.get_xfid( p_alignment => as_xlsx.get_alignment( p_horizontal => 'centerContinuous' ) )
             );
  as_xlsx.set_column_width( p_col   => 5
                          , p_width => 15
                           );
  as_xlsx.set_column_width( p_col   => 6
                          , p_width => 15
                          );
  as_xlsx.cell( 5
              , l_cnt
                 + 3  -- query start row
                 + 2  -- title + headers 
                 + 1  -- interval 
              , 'Rows returned: ' || l_cnt );
  -- make sure you have set as_xlsx.use_dbms_crypto = true; in the package specification
  as_xlsx.save( 'EXCEL_OUT', 'anton-tables.xlsx');
end;
0
0
