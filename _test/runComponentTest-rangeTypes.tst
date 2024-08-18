PL/SQL Developer Test script 3.0
54
DECLARE

   TYPE tp_cell_loc IS RECORD (
      c     PLS_INTEGER, -- 2
      r     PLS_INTEGER, -- 3
      fix1  BOOLEAN := false,
      fix2  BOOLEAN := false, -- true
      alfan VARCHAR2(12) );  -- B$3
   TYPE tp_cell_range IS RECORD (
      tl    tp_cell_loc,
      br    tp_cell_loc,
      alfan VARCHAR2(12) );
   rg_  tp_cell_range;

   FUNCTION Alfan_Col (
      col_ IN PLS_INTEGER ) RETURN VARCHAR2
   IS BEGIN
      RETURN CASE
         WHEN col_ > 702 THEN chr(64+trunc((col_-27)/676)) || chr(65+mod(trunc((col_-1)/26)-1, 26)) || chr(65+mod(col_-1, 26))
         WHEN col_ > 26  THEN chr(64+trunc((col_-1)/26)) || chr(65+mod(col_-1, 26))
         ELSE chr(64+col_)
      END;
   END Alfan_Col;

   FUNCTION Alfan_Cell (
      col_  IN PLS_INTEGER,
      row_  IN PLS_INTEGER,
      fix1_ IN BOOLEAN := false,
      fix2_ IN BOOLEAN := false ) RETURN VARCHAR2
   IS
      d1_  VARCHAR2(1) := CASE WHEN fix1_ THEN '$' END;
      d2_  VARCHAR2(1) := CASE WHEN fix2_ THEN '$' END;
   BEGIN
      RETURN d1_ || Alfan_Col(col_) || d2_ || to_char(row_);
   END Alfan_Cell;

   FUNCTION To_Alfan (
      loc_ IN OUT NOCOPY tp_cell_loc ) RETURN VARCHAR2
   IS BEGIN
      loc_.alfan := Alfan_Cell (loc_.c, loc_.r, loc_.fix1, loc_.fix2);
      RETURN loc_.alfan;
   END To_Alfan;
   PROCEDURE To_Alfan (
      range_ IN OUT NOCOPY tp_cell_range )
   IS
   BEGIN
      range_.alfan := To_Alfan (range_.tl) || ':' || To_Alfan (range_.br);
   END To_Alfan;
BEGIN
   rg_ := tp_cell_range (tp_cell_loc(2,3,false,true),tp_cell_loc(6,6));
   Dbms_Output.Put_Line ('alfan cell is: ' || nvl(rg_.alfan, '<< null >>'));
   To_Alfan (rg_);
   Dbms_Output.Put_Line ('alfan cell is: ' || nvl(rg_.alfan, '<< null >>'));
END;
0
0
