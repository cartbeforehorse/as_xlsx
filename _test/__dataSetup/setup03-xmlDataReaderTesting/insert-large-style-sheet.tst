PL/SQL Developer Test script 3.0
503
DECLARE
   val1_ CLOB := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <numFmts count="11">
    <numFmt numFmtId="166" formatCode="#,##0"/>
    <numFmt numFmtId="167" formatCode="#,##0.00"/>
    <numFmt numFmtId="174" formatCode="Mmm yyyy"/>
    <numFmt numFmtId="165" formatCode="_-&quot;£&quot;* #,##0.00_-;-&quot;£&quot;* #,##0.00_-;_-&quot;£&quot;* &quot;-&quot;_-;_-@_-"/>
    <numFmt numFmtId="164" formatCode="_-&quot;£&quot;* #,##0_-;-&quot;£&quot;* #,##0_-;_-&quot;£&quot;* &quot;-&quot;_-;_-@_-"/>
    <numFmt numFmtId="168" formatCode="dd mmm yyyy"/>
    <numFmt numFmtId="169" formatCode="dd mmm yyyy hh:mm"/>
    <numFmt numFmtId="170" formatCode="dd mmm yyyy hh:mm AM/PM"/>
    <numFmt numFmtId="171" formatCode="dd mmm yyyy hh:mm:ss"/>
    <numFmt numFmtId="172" formatCode="dd mmm yyyy hh:mm:ss AM/PM"/>
    <numFmt numFmtId="173" formatCode="dd mmmm yyyy"/>
  </numFmts>
  <fonts count="13" x14ac:knownFonts="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFDBE5F1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FF244062"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFDCE6F1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FFDCE6F1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFEBF1DE"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <i/>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <sz val="11"/>
      <color rgb="FF4F6228"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
    <font>
      <u/>
      <sz val="11"/>
      <color theme="10"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
  </fonts>';
  val2_ CLOB := '  <fills count="15">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF17375D"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF366092"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF95B3D7"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF953735"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF006400"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFD8E4BC"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF76933C"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFDCE6F1"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF60497A"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFF2F2F2"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFD9D9D9"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFA6A6A6"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF595959"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="46">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="dotted"/>
      <right style="dotted"/>
      <top style="dotted"/>
      <bottom style="dotted"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="dotted"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="dotted"/>
      <right style="none"/>
      <top style="dotted"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="dotted"/>
      <right style="none"/>
      <top style="dotted"/>
      <bottom style="dotted"/>
    </border>
    <border>
      <left style="none"/>
      <right style="dotted"/>
      <top style="dotted"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="dotted"/>
      <bottom style="dotted"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="dotted"/>
    </border>
    <border>
      <left style="dotted"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="dotted"/>
    </border>
    <border>
      <left style="dotted"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="dotted"/>
      <top style="none"/>
      <bottom style="dotted"/>
    </border>
    <border>
      <left style="none"/>
      <right style="dotted"/>
      <top style="none"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="dotted"/>
      <top style="dotted"/>
      <bottom style="dotted"/>
    </border>
    <border>
      <left style="thin"/>
      <right style="thin"/>
      <top style="thin"/>
      <bottom style="thin"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="thin"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="thin"/>
      <right style="none"/>
      <top style="thin"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="thin"/>
      <right style="none"/>
      <top style="thin"/>
      <bottom style="thin"/>
    </border>
    <border>
      <left style="none"/>
      <right style="thin"/>
      <top style="thin"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="thin"/>
      <bottom style="thin"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="thin"/>
    </border>
    <border>
      <left style="thin"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="thin"/>
    </border>
    <border>
      <left style="thin"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="thin"/>
      <top style="none"/>
      <bottom style="thin"/>
    </border>
    <border>
      <left style="none"/>
      <right style="thin"/>
      <top style="none"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="thin"/>
      <top style="thin"/>
      <bottom style="thin"/>
    </border>
    <border>
      <left style="medium"/>
      <right style="medium"/>
      <top style="medium"/>
      <bottom style="medium"/>
    </border>';
   val3_ CLOB := '    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="medium"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="medium"/>
      <right style="none"/>
      <top style="medium"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="medium"/>
      <right style="none"/>
      <top style="medium"/>
      <bottom style="medium"/>
    </border>
    <border>
      <left style="none"/>
      <right style="medium"/>
      <top style="medium"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="medium"/>
      <bottom style="medium"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="medium"/>
    </border>
    <border>
      <left style="medium"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="medium"/>
    </border>
    <border>
      <left style="medium"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="medium"/>
      <top style="none"/>
      <bottom style="medium"/>
    </border>
    <border>
      <left style="none"/>
      <right style="medium"/>
      <top style="none"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="medium"/>
      <top style="medium"/>
      <bottom style="medium"/>
    </border>
    <border>
      <left style="thick"/>
      <right style="thick"/>
      <top style="thick"/>
      <bottom style="thick"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="thick"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="thick"/>
      <right style="none"/>
      <top style="thick"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="thick"/>
      <top style="thick"/>
      <bottom style="none"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="thick"/>
      <bottom style="thick"/>
    </border>
    <border>
      <left style="none"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="thick"/>
    </border>
    <border>
      <left style="thick"/>
      <right style="none"/>
      <top style="none"/>
      <bottom style="thick"/>
    </border>
    <border>
      <left style="none"/>
      <right style="thick"/>
      <top style="none"/>
      <bottom style="thick"/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="" fontId="1" fillId="2" borderId=""/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="10" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="8" fillId="3" borderId="0"/>
    <xf numFmtId="0" fontId="12" fillId="0" borderId="0"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="167" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="171" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="174" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
  <extLst>
    <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
    </ext>
  </extLst>
</styleSheet>';
BEGIN
   DELETE FROM xml_data_tester WHERE id = 5;
   INSERT INTO xml_data_tester dt (dt.id, dt.description, dt.xml_content)
   VALUES (5, 'Styles part, but more exciting', val1_ || val2_ || val3_);
   commit;
END;
0
0
