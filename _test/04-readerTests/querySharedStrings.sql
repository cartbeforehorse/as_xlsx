SELECT xt1.txt, xt1.len
FROM   xml_data_tester dt
       CROSS JOIN xmlTable (
          xmlNamespaces (
             default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
             'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x"
          ), '(/sst/si, /x:sst/x:si)'
          passing xmltype (dt.xml_content)
          columns txt VARCHAR2(4000 CHAR) path 'substring(string-join(.//*:t/text(),""), 1, 3900)',
                  len INTEGER             path 'string-length(string-join(.//*:t/text(), ""))'
       ) xt1
