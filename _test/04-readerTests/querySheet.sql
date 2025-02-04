SELECT dt.id, dt.description, Dbms_Lob.Substr (dt.xml_content, 3700) style_xml,
       xt1.*
FROM   xml_data_tester dt
       CROSS JOIN xmlTable (
          xmlNamespaces (
             default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                   'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x"
                ), '(/worksheet/sheetData/row/c, /x:worksheet/x:sheetData/x:row/x:c)'
          passing xmltype(dt.xml_content)
             columns c_val   VARCHAR2(4000) path '*:v',
                     f       VARCHAR2(4000) path '*:f',
                     c_type  VARCHAR2(4000) path '@t',
                     c_ref   VARCHAR2(32)   path '@r',
                     c_style INTEGER        path '@s',
                     c_row   INTEGER        path './../@r',
                     txt VARCHAR2(4000 CHAR) path 'substring(string-join(.//*:t/text(), ""), 1, 3900)',
                     len INTEGER        path 'string-length(string-join(.//*:t/text(),""))'
       ) xt1
WHERE  dt.id = 7
