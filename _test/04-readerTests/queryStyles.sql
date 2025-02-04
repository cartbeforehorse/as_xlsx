SELECT dt.id, dt.description, Dbms_Lob.Substr (dt.xml_content, 3700) style_xml,
       xt2.seq - 1 seq, xt2.id, xt2.quoteprefix, lower(xt3.format) format
FROM   xml_data_tester dt
       CROSS JOIN xmlTable (
          xmlNamespaces (
             default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
             'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x"
          ), '(/styleSheet, /x:styleSheet)'
          passing xmltype(dt.xml_content)
          columns cellxfs xmltype path '(cellXfs, x:cellXfs)',
                  numfmts xmltype path '(numFmts, x:numFmts)'
       ) xt1
       CROSS JOIN xmlTable (
          '/*:cellXfs/*:xf'
          passing xt1.cellxfs
          columns seq FOR ordinality,
                  id          INTEGER        path '@numFmtId',
                  quoteprefix VARCHAR2(4000) path '@quotePrefix'
       ) xt2
       LEFT JOIN xmlTable (
          '/*:numFmts/*:numFmt'
          passing xt1.numfmts
          columns id3    INTEGER        path '@numFmtId',
                  format VARCHAR2(4000) path '@formatCode'
       ) xt3
         ON xt3.id3 = xt2.id
WHERE  dt.id = 5
