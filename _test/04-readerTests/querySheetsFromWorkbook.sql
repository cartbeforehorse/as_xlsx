WITH file_parts AS (
   SELECT (SELECT dt.xml_content FROM xml_data_tester dt WHERE dt.id = 2) wb,
          (SELECT dt.xml_content FROM xml_data_tester dt WHERE dt.id = 6) wbr
   FROM   dual
)
SELECT Dbms_Lob.Substr(fp.wb,4000) wb, Dbms_Lob.Substr(fp.wbr,4000) wbr,
       xt1.d1904, xt2.seq, xt2.sheet_name, xt2.sheet_id, xt2.rid, xt2.state, xt3.type, xt3.target, xt3.id
FROM   file_parts fp
       CROSS JOIN xmlTable (
          xmlNamespaces (
             default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
             'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x"
          ), '(/workbook, /x:workbook)'
          passing xmltype (fp.wb)
          columns d1904  VARCHAR2(4000) path '*:workbookPr/@date1904',
                  sheets xmltype        path '*:sheets'
       ) xt1
       CROSS JOIN xmlTable (
          '*:sheets/*:sheet' passing xt1.sheets
          columns seq for ordinality,
                  sheet_name VARCHAR2(4000) path '@name',
                  sheet_id   VARCHAR2(4000) path '@sheetId',
                  rid        VARCHAR2(4000) path '@*:id',
                  state      VARCHAR2(4000) path '@state'
       ) xt2
       INNER JOIN xmlTable (
          xmlNamespaces (default 'http://schemas.openxmlformats.org/package/2006/relationships'),
          '/Relationships/Relationship'
          passing xmlType (fp.wbr)
          columns type   VARCHAR2(4000)   path '@Type',
                  target VARCHAR2(2000) path '@Target',
                  id     VARCHAR2(2000) path '@Id'
       ) xt3
          ON xt3.id = xt2.rid
