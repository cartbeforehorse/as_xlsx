INSERT INTO xml_data_tester dt (dt.id, dt.description, dt.xml_content)
VALUES (1, 'Excel "workbook.xml" example 01', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
  <workbookPr defaultThemeVersion="166925" date1904="false"/>
  <bookViews>
    <workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
  </bookViews>
  <sheets>
    <sheet name="Data and Tables" sheetId="1" r:id="rId4"/>
  </sheets>
  <calcPr calcId="144525"/>
</workbook>');
INSERT INTO xml_data_tester dt (dt.id, dt.description, dt.xml_content)
VALUES (2, 'Multi page example', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
  <workbookPr defaultThemeVersion="166925" date1904="false"/>
  <bookViews>
    <workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
  </bookViews>
  <sheets>
    <sheet name="Parameters" sheetId="1" r:id="rId4"/>
    <sheet name="Number Two" sheetId="2" r:id="rId5"/>
    <sheet name="Data" sheetId="3" r:id="rId6"/>
    <sheet name="Number Four" sheetId="4" r:id="rId7"/>
  </sheets>
  <definedNames>
    <definedName name="CustomerData">&apos;Parameters&apos;!$B$10:$C$13</definedName>
    <definedName name="MyDataSource">&apos;Data&apos;!$B$3:$E$13</definedName>
  </definedNames>
  <calcPr calcId="144525"/>
</workbook>');
INSERT INTO xml_data_tester dt (dt.id, dt.description, dt.xml_content)
VALUES (3, 'Shared Strings part', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="154" uniqueCount="25">
  <si><t xml:space="preserve">Company</t></si>
  <si><t xml:space="preserve">Identity Type</t></si>
  <si><t xml:space="preserve">Identity</t></si>
  <si><t xml:space="preserve">Category</t></si>
  <si><t xml:space="preserve">Currency</t></si>
  <si><t xml:space="preserve">Amount</t></si>
  <si><t xml:space="preserve">Tax</t></si>
  <si><t xml:space="preserve">HOLD</t></si>
  <si><t xml:space="preserve">Supplier</t></si>
  <si><t xml:space="preserve">Supp Dude</t></si>
  <si><t xml:space="preserve">Brakes</t></si>
  <si><t xml:space="preserve">EUR</t></si>
  <si><t xml:space="preserve">LTD</t></si>
  <si><t xml:space="preserve">George Doors</t></si>
  <si><t xml:space="preserve">Gears</t></si>
  <si><t xml:space="preserve">GBP</t></si>
  <si><t xml:space="preserve">Customer</t></si>
  <si><t xml:space="preserve">Car Bits</t></si>
  <si><t xml:space="preserve">Green Engines</t></si>
  <si><t xml:space="preserve">Ford</t></si>
  <si><t xml:space="preserve">Interior</t></si>
  <si><t xml:space="preserve">SEK</t></si>
  <si><t xml:space="preserve">Volvo</t></si>
  <si><t xml:space="preserve">USD</t></si>
  <si><t xml:space="preserve">LLB</t></si>
</sst>');
INSERT INTO xml_data_tester dt (dt.id, dt.description, dt.xml_content)
VALUES (4, 'Styles part', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <fonts count="1" x14ac:knownFonts="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="none"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
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
</styleSheet>');

INSERT INTO xml_data_tester dt (dt.id, dt.description, dt.xml_content)
VALUES (6, '"workbook.xml.rels" that matches ID:2', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
</Relationships>');

INSERT INTO xml_data_tester dt (dt.id, dt.description, dt.xml_content)
VALUES (7, '"sheet1.xml" from a one-sheet workbook', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="x14ac" xr:uid="{2D425352-AFD0-0193-E063-020011ACA923}">
  <dimension ref="B2:C11"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="A1" sqref="A1"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
  <cols>
    <col min="3" max="3" width="20.7109375" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="2" spans="2:3">
      <c r="B2" t="s">
        <v>0</v>
      </c>
    </row>
    <row r="3" spans="2:3">
      <c r="B3" t="s">
        <v>1</v>
      </c>
      <c r="C3" t="s" s="2">
        <v>2</v>
      </c>
    </row>
    <row r="4" spans="2:3">
      <c r="B4" t="s">
        <v>3</v>
      </c>
      <c r="C4" t="s" s="3">
        <v>4</v>
      </c>
    </row>
    <row r="5" spans="2:3">
      <c r="B5" t="s">
        <v>5</v>
      </c>
      <c r="C5" t="s" s="4">
        <v>6</v>
      </c>
    </row>
    <row r="6" spans="2:3">
      <c r="B6" t="s">
        <v>7</v>
      </c>
      <c r="C6" t="s" s="5">
        <v>8</v>
      </c>
    </row>
    <row r="7" spans="2:3">
      <c r="B7" t="s">
        <v>9</v>
      </c>
      <c r="C7" s="6">
        <v>123.657</v>
      </c>
    </row>
    <row r="8" spans="2:3">
      <c r="B8" t="s">
        <v>10</v>
      </c>
      <c r="C8" s="7">
        <v>.56</v>
      </c>
    </row>
    <row r="9" spans="2:3">
      <c r="B9" t="s">
        <v>11</v>
      </c>
      <c r="C9">
        <v>43563.9899665367</v>
      </c>
    </row>
    <row r="10" spans="2:3">
      <c r="B10" t="s">
        <v>12</v>
      </c>
      <c r="C10" s="8">
        <v>45292.5738541666666666666666666666666667</v>
      </c>
    </row>
    <row r="11" spans="2:3">
      <c r="B11" t="s">
        <v>13</v>
      </c>
      <c r="C11" s="9">
        <v>45292.5738541666666666666666666666666667</v>
      </c>
    </row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="C6" r:id="rId1"/>
  </hyperlinks>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <legacyDrawing r:id="rId2"/>
</worksheet>');
