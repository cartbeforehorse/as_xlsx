CREATE OR REPLACE PACKAGE nyce_xml IS

   -----
   -- Datatype definitions
   --   Datatypes that help with managing attributes on XML nodes
   --
   TYPE xml_attr IS RECORD (
      key VARCHAR2(200),
      val VARCHAR2(2000)
   );
   TYPE xml_attrs_array IS TABLE OF xml_attr INDEX BY PLS_INTEGER;
   TYPE attrs_pos IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(200);
   TYPE xml_attrs_arr IS RECORD (
      attrs    xml_attrs_array,
      attr_loc attrs_pos
   );

   ---
   -- Author  : Now you can, ey! (NYCE)
   -- Created : 2024-10-27
   -- Purpose :
   --   A package to help the application build XML files
   -- 
   -- Developer Notes:
   -- Please observe coding conveitions:
   --   - 3-space indents
   --   - uppercase oracle code-structure keywords
   --   - camel-case underscored functions/procedure names
   --   - lowercase everything else
   --   - trailing underscore for variables (none of this v_, p_ silliness)
   --   - Stacked variables to appear indented beneath function calls, and not
   --     aligned with the end of their function name
   --   - try to write your code in well indented "blocks" of code so that the
   --     next person has a chance of understanding your thoughts!  Often this
   --     will be you debugging your own mistakes, so be nyce!!
   --   - !!! NEVER USE CODE BAUTIFIERS !!!  They make code look ugly and mess
   --     up version control!
   --

   ---------------------------------------
   -- Supporting functions...
   --
   FUNCTION Clob_To_Blob (
      clob_in_ IN CLOB ) RETURN BLOB;

   ---------------------------------------
   -- XML attributes management
   --
   PROCEDURE cAtr (
      attrs_ IN OUT NOCOPY xml_attrs_arr );
   PROCEDURE Attr (
      key_   IN VARCHAR2,
      val_   IN VARCHAR2,
      attrs_ IN OUT NOCOPY xml_attrs_arr );
   PROCEDURE nAtr (
      key_   IN VARCHAR2,
      val_   IN VARCHAR2,
      attrs_ IN OUT NOCOPY xml_attrs_arr );

   ---------------------------------------
   -- XML file "part" builders
   --
   FUNCTION Xml_Date_Time (
      dt_ IN DATE ) RETURN VARCHAR2;
   FUNCTION Xml_Date_Time_Tz (
      dt_ IN DATE ) RETURN VARCHAR2;
   FUNCTION Xml_Date (
      dt_ IN DATE ) RETURN VARCHAR2;
   FUNCTION Xml_Time_Mi (
      dt_ IN DATE ) RETURN VARCHAR2;
   FUNCTION Xml_Number (
      num_ IN NUMBER,
      fm_  IN VARCHAR2 := null ) RETURN VARCHAR2;
   FUNCTION Xml_Number (
      num_      IN NUMBER,
      decimals_ IN NUMBER := 5 ) RETURN VARCHAR2;
   FUNCTION Make_Node (
      doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
      tag_name_ IN VARCHAR2,
      ns_       IN VARCHAR2      := '',
      attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode;
   PROCEDURE Make_Node (
      doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
      tag_name_ IN VARCHAR2,
      ns_       IN VARCHAR2      := '',
      attrs_    IN xml_attrs_arr := cast (null as xml_attrs_arr) );

   ---------------------------------------
   -- XML DOM Builders
   --
   FUNCTION Make_Root_Node (
      doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
      tag_name_ IN VARCHAR2,
      ns_       IN VARCHAR2 := '',
      xsd_loc_  IN VARCHAR2 := '' ) RETURN dbms_XmlDom.DomNode;
   FUNCTION Make_Root_Node (
      doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
      tag_name_ IN VARCHAR2,
      ns_       IN VARCHAR2      := '',
      attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode;


   FUNCTION Xml_Node (
      doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_ IN dbms_XmlDom.DomNode,
      tag_name_  IN VARCHAR2,
      ns_        IN VARCHAR2,
      attrs_     IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode;
   FUNCTION Xml_Node (
      doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_ IN dbms_XmlDom.DomNode,
      tag_name_  IN VARCHAR2,
      attrs_     IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode;
   PROCEDURE Xml_Node (
      doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_ IN dbms_XmlDom.DomNode,
      tag_name_  IN VARCHAR2,
      ns_        IN VARCHAR2,
      attrs_     IN xml_attrs_arr := xml_attrs_arr() );
   PROCEDURE Xml_Node (
      doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_ IN dbms_XmlDom.DomNode,
      tag_name_  IN VARCHAR2,
      attrs_     IN xml_attrs_arr := xml_attrs_arr() );

   PROCEDURE Xml_Text_Node (
      doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_    IN dbms_XmlDom.DomNode,
      tag_name_     IN VARCHAR2,
      text_content_ IN VARCHAR2,
      ns_           IN VARCHAR2,
      attrs_        IN xml_attrs_arr := xml_attrs_arr() );
   PROCEDURE Xml_Text_Node (
      doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_    IN dbms_XmlDom.DomNode,
      tag_name_     IN VARCHAR2,
      text_content_ IN VARCHAR2,
      attrs_        IN xml_attrs_arr := xml_attrs_arr() );
   PROCEDURE Xml_Text_Node (
      doc_         IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_   IN dbms_XmlDom.DomNode,
      tag_name_    IN VARCHAR2,
      num_content_ IN NUMBER,
      decimals_    IN NUMBER        := 0,
      ns_          IN VARCHAR2,
      attrs_       IN xml_attrs_arr := xml_attrs_arr() );
   PROCEDURE Xml_Text_Node (
      doc_         IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_   IN dbms_XmlDom.DomNode,
      tag_name_    IN VARCHAR2,
      num_content_ IN NUMBER,
      decimals_    IN NUMBER        := 0,
      attrs_       IN xml_attrs_arr := xml_attrs_arr() );

   PROCEDURE Xml_Clob_Node (
      doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_    IN dbms_XmlDom.DomNode,
      tag_name_     IN VARCHAR2,
      clob_content_ IN CLOB,
      ns_           IN VARCHAR2      := '',
      attrs_        IN xml_attrs_arr := xml_attrs_arr() );
   PROCEDURE Xml_Blob_Node (
      doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
      append_to_    IN dbms_XmlDom.DomNode,
      tag_name_     IN VARCHAR2,
      blob_content_ IN BLOB,
      ns_           IN VARCHAR2      := '',
      attrs_        IN xml_attrs_arr := xml_attrs_arr() );

   ------------
   -- Silly Foundation stuff
   PROCEDURE Init;

END nyce_xml;
/
CREATE OR REPLACE PACKAGE BODY nyce_xml IS

-----
--   The following code-fragments are built as wrappers around the excessively
--   complex package Dbms_XmlDom.  They simplify the process of generating XML
--   documents by burying much of the complexity into simpler functions.  It's
--   also quite nice to have PROCEDURE versions of the XML fragment generators
--   because we don't always need to store a reference to a newly created node
--   or element.

-----
-- LOB support functions
--
FUNCTION Clob_To_Blob (
   clob_in_ IN CLOB ) RETURN BLOB
IS
   blob_out_     BLOB;
   dest_offset_  INTEGER := 1;
   src_offset_   INTEGER := 1;
   lang_context_ INTEGER := dbms_lob.DEFAULT_LANG_CTX;
   warning_      INTEGER;
BEGIN
   Dbms_Lob.CreateTemporary (blob_out_, true);
   Dbms_Lob.ConvertToBlob (
      dest_lob     => blob_out_,
      src_clob     => clob_in_,
      amount       => length(clob_in_),
      dest_offset  => dest_offset_,
      src_offset   => src_offset_,
      blob_csid    => nls_charset_id('AL32UTF8'),
      lang_context => lang_context_,
      warning      => warning_
   );
   RETURN blob_out_;   
END Clob_To_Blob;

FUNCTION Blob_To_Clob64 (
   blob_ IN BLOB ) RETURN CLOB
IS
   offset_         NUMBER := 1;
   chunk_size_     NUMBER := 1902;
   blob_chunk_     RAW(1902);
   data_as_base64_ CLOB;
BEGIN
   WHILE offset_ <= Dbms_Lob.GetLength(blob_) LOOP
      Dbms_Lob.Read (blob_, chunk_size_, offset_, blob_chunk_);
      data_as_base64_ := data_as_base64_ || Utl_Encode.Base64_Encode (blob_chunk_);
      offset_ := offset_ + chunk_size_;
   END LOOP;
   RETURN data_as_base64_;
END Blob_To_Clob64;


---------------------------------------
-- XML attributes management
--
PROCEDURE cAtr (
   attrs_ IN OUT NOCOPY xml_attrs_arr )
IS BEGIN
   attrs_.attrs.delete;
   attrs_.attr_loc.delete;
END cAtr;

PROCEDURE Attr (
   key_   IN VARCHAR2,
   val_   IN VARCHAR2,
   attrs_ IN OUT NOCOPY xml_attrs_arr )
IS
   ix_ PLS_INTEGER := attrs_.attrs.count + 1;
BEGIN
   IF not attrs_.attr_loc.exists(key_) THEN
      ix_ := attrs_.attrs.count + 1;
      attrs_.attr_loc(key_) := ix_;
      attrs_.attrs(ix_).key := key_;
   ELSE
      ix_ := attrs_.attr_loc(key_);
   END IF;
   attrs_.attrs(ix_).val := val_;
END Attr;

PROCEDURE nAtr (
   key_   IN VARCHAR2,
   val_   IN VARCHAR2,
   attrs_ IN OUT NOCOPY xml_attrs_arr )
IS BEGIN
   cAtr (attrs_);
   Attr (key_, val_, attrs_);
END nAtr;


---------------------------------------
-- XML file "part" builders
--
FUNCTION Xml_Date_Time (
   dt_ IN DATE ) RETURN VARCHAR2
IS BEGIN
   RETURN replace (to_char(dt_, 'YYYY-MM-DD_HH24:MI:SS'),'_','T');
END Xml_Date_Time;

FUNCTION Xml_Date_Time_Tz (
   dt_ IN DATE ) RETURN VARCHAR2
IS BEGIN
   RETURN Xml_Date_Time (dt_) || sessiontimezone;
END Xml_Date_Time_Tz;

FUNCTION Xml_Date (
   dt_ IN DATE ) RETURN VARCHAR2
IS BEGIN
   RETURN to_char(dt_, 'YYYY-MM-DD');
END Xml_Date;

FUNCTION Xml_Time_Mi (
   dt_ IN DATE ) RETURN VARCHAR2
IS BEGIN
   RETURN to_char(dt_, 'HH24:MI');
END Xml_Time_Mi;

FUNCTION Xml_Number (
   num_ IN NUMBER,
   fm_  IN VARCHAR2 := null ) RETURN VARCHAR2
IS
   mask_ VARCHAR2(99) := nvl (fm_, 'fm99999999999999999999.99999');
BEGIN
   RETURN rtrim (to_char (num_, mask_), '.');
END Xml_Number;

FUNCTION Xml_Number (
   num_      IN NUMBER,
   decimals_ IN NUMBER := 5 ) RETURN VARCHAR2
IS
   nr_decimals_  NUMBER        := decimals_+1;
   decimal_mask_ VARCHAR2(20)  := rpad ('.', nr_decimals_, '9');
   mask_         VARCHAR2(99)  := 'fm999999999999999999999999999' || decimal_mask_;
BEGIN
   RETURN rtrim (to_char (num_, mask_), '.');
END Xml_Number;

FUNCTION Make_Tag (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomElement
IS
   el_ dbms_XmlDom.DomElement := CASE
      WHEN ns_ IS NOT null THEN Dbms_XmlDom.createElement (doc_, tag_name_, ns_)
      ELSE Dbms_XmlDom.createElement (doc_, tag_name_)
   END;
BEGIN
   FOR ix_ IN 1 .. attrs_.attrs.count LOOP
      Dbms_XmlDom.setAttribute (el_, attrs_.attrs(ix_).key, attrs_.attrs(ix_).val);
   END LOOP;
   RETURN el_;
END Make_Tag;

---------------------------------------
-- XML DOM Builders
--
FUNCTION Make_Root_Node (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2 := '',
   xsd_loc_  IN VARCHAR2 := '' ) RETURN dbms_XmlDom.DomNode
IS
   attrs_ xml_attrs_arr;
BEGIN
   IF ns_ IS NOT null THEN
      attr ('xmlns', ns_, attrs_);
   END IF;
   IF xsd_loc_ IS NOT null THEN
      attr ('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', attrs_);
      attr ('xsi:schemaLocation', xsd_loc_, attrs_);
   END IF;
   RETURN Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), tag_name_, ns_, attrs_);
END Make_Root_Node;

FUNCTION Make_Root_Node (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode
IS BEGIN
   RETURN Xml_Node (doc_, Dbms_XmlDom.makeNode(doc_), tag_name_, ns_, attrs_);
END Make_Root_Node;

FUNCTION Make_Node (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode
IS
   nd_ dbms_XmlDom.DomNode := Dbms_XmlDom.makeNode (Make_Tag (doc_, tag_name_, ns_, attrs_));
BEGIN
   IF ns_ IS NOT null THEN
      Dbms_XmlDom.setPrefix (nd_, ns_);
   END IF;
   RETURN nd_;
END Make_Node;

PROCEDURE Make_Node (
   doc_      IN OUT NOCOPY dbms_XmlDom.DomDocument,
   tag_name_ IN VARCHAR2,
   ns_       IN VARCHAR2      := '',
   attrs_    IN xml_attrs_arr := cast (null as xml_attrs_arr) )
IS
   throw_nd_ dbms_XmlDom.DomNode;
BEGIN
   throw_nd_ := Make_Node (doc_, tag_name_, ns_, attrs_);
END Make_Node;

FUNCTION Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   ns_        IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode
IS
   nd_  dbms_XmlDom.DomNode := append_to_;
   CURSOR tag_hierarchy IS
      SELECT regexp_substr (tag_name_, '[^/]+', 1, level) tag_name, level lv,
             count(1) over() nr
      FROM   dual
      CONNECT BY regexp_substr (tag_name_, '[^/]+', 1, level) IS NOT null
      ORDER BY lv ASC;
BEGIN
   FOR h_ IN tag_hierarchy LOOP
      IF h_.lv = h_.nr THEN
         nd_ := Dbms_XmlDom.appendChild (nd_, Make_Node(doc_,h_.tag_name,ns_,attrs_));
      ELSE
         nd_ := Dbms_XmlDom.appendChild (nd_, Make_Node(doc_,h_.tag_name,ns_));
      END IF;
   END LOOP;
   RETURN nd_;
END Xml_Node;

FUNCTION Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() ) RETURN dbms_XmlDom.DomNode
IS BEGIN
   RETURN Xml_Node (doc_, append_to_, tag_name_, '', attrs_);
END Xml_Node;

PROCEDURE Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   ns_        IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() )
IS
   throw_nd_ dbms_XmlDom.DomNode;
BEGIN
   throw_nd_ := Xml_Node (doc_, append_to_, tag_name_, ns_, attrs_);
END Xml_Node;

PROCEDURE Xml_Node (
   doc_       IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_ IN dbms_XmlDom.DomNode,
   tag_name_  IN VARCHAR2,
   attrs_     IN xml_attrs_arr := xml_attrs_arr() )
IS
   throw_nd_ dbms_XmlDom.DomNode;
BEGIN
   throw_nd_ := Xml_Node (doc_, append_to_, tag_name_, '', attrs_);
END Xml_Node;

PROCEDURE Xml_Text_Node (
   doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_    IN dbms_XmlDom.DomNode,
   tag_name_     IN VARCHAR2,
   text_content_ IN VARCHAR2,
   ns_           IN VARCHAR2,
   attrs_        IN xml_attrs_arr := xml_attrs_arr() )
IS
   nd_ dbms_XmlDom.DomNode := append_to_;
   CURSOR tag_hierarchy IS
      SELECT regexp_substr (tag_name_, '[^/]+', 1, level) tag_name, level lv,
             count(1) over() nr
      FROM   dual
      CONNECT BY regexp_substr (tag_name_, '[^/]+', 1, level) IS NOT null
      ORDER BY lv ASC;
BEGIN
   FOR h_ IN tag_hierarchy LOOP
      IF h_.lv = h_.nr THEN
         nd_ := Dbms_XmlDom.appendChild (
            Dbms_XmlDom.appendChild (nd_, Make_Node(doc_,h_.tag_name,ns_,attrs_)),
            Dbms_XmlDom.makeNode (Dbms_XmlDom.createTextNode (doc_, text_content_))
         );
      ELSE
         nd_ := Dbms_XmlDom.appendChild (nd_, Make_Node(doc_,h_.tag_name,ns_));
      END IF;
   END LOOP;
END Xml_Text_Node;

PROCEDURE Xml_Text_Node (
   doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_    IN dbms_XmlDom.DomNode,
   tag_name_     IN VARCHAR2,
   text_content_ IN VARCHAR2,
   attrs_        IN xml_attrs_arr := xml_attrs_arr() )
IS BEGIN
   Xml_Text_Node (doc_, append_to_, tag_name_, text_content_, '', attrs_);
END Xml_Text_Node;

PROCEDURE Xml_Text_Node (
   doc_         IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_   IN dbms_XmlDom.DomNode,
   tag_name_    IN VARCHAR2,
   num_content_ IN NUMBER,
   decimals_    IN NUMBER        := 0,
   ns_          IN VARCHAR2,
   attrs_       IN xml_attrs_arr := xml_attrs_arr() )
IS BEGIN
   Xml_Text_Node (
      doc_          => doc_,
      append_to_    => append_to_,
      tag_name_     => tag_name_,
      text_content_ => Xml_Number (num_content_, decimals_),
      ns_           => ns_,
      attrs_        => attrs_
   );
END Xml_Text_Node;

PROCEDURE Xml_Text_Node (
   doc_         IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_   IN dbms_XmlDom.DomNode,
   tag_name_    IN VARCHAR2,
   num_content_ IN NUMBER,
   decimals_    IN NUMBER        := 0,
   attrs_       IN xml_attrs_arr := xml_attrs_arr() )
IS BEGIN
   Xml_Text_Node (doc_, append_to_, tag_name_, num_content_, decimals_, '', attrs_);
END Xml_Text_Node;

PROCEDURE Xml_Clob_Node (
   doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_    IN dbms_XmlDom.DomNode,
   tag_name_     IN VARCHAR2,
   clob_content_ IN CLOB,
   ns_           IN VARCHAR2      := '',
   attrs_        IN xml_attrs_arr := xml_attrs_arr() )
IS
   chunk_      VARCHAR2(4000);
   chunk_size_ NUMBER                        := 4000;
   offset_     NUMBER                        := 1;
   txt_el_     Dbms_XmlDom.DomText           := Dbms_XmlDom.createTextNode (doc_, '');
   nd_txt_     Dbms_XmlDom.DomNode           := Dbms_XmlDom.makeNode (txt_el_);
   stream_     sys.utl_characterOutputStream := Dbms_XmlDom.setNodeValueAsCharacterStream (nd_txt_);
   throw_nd_   dbms_XmlDom.DomNode;
BEGIN
   WHILE offset_ <= Dbms_Lob.GetLength(clob_content_) LOOP
      Dbms_Lob.Read (clob_content_, chunk_size_, offset_, chunk_);
      stream_.write (chunk_, 0, chunk_size_);
      offset_ := offset_ + chunk_size_;
   END LOOP;
   stream_.close;
   throw_nd_ := Dbms_XmlDom.appendChild (
      Dbms_XmlDom.appendChild (append_to_, Make_Node(doc_,tag_name_,ns_,attrs_)),
      nd_txt_
   );
END Xml_Clob_Node;

-----
-- Add_Blob_Node_To_Dom()
--   Remember that converting binary into base64 increases storage consumption
--   by 25%.  3 BLOB bytes are converted into 4 CLOB bytes.  Also, I pick 2496
--   because it divides by both 3 and 64 (Utl_Raw.Cast_To_Varchar2() appends a
--   new-line character every 64 characters.
PROCEDURE Xml_Blob_Node (
   doc_          IN OUT NOCOPY dbms_XmlDom.DomDocument,
   append_to_    IN dbms_XmlDom.DomNode,
   tag_name_     IN VARCHAR2,
   blob_content_ IN BLOB,
   ns_           IN VARCHAR2      := '',
   attrs_        IN xml_attrs_arr := xml_attrs_arr() )
IS
   chunk_base64_ VARCHAR2(32000);
   chunk64_len_  NUMBER;
   blob_chunk_   RAW(2496);
   chunk_size_   NUMBER                        := 2496;
   offset_       NUMBER                        := 1;
   txt_el_       Dbms_XmlDom.DomText           := Dbms_XmlDom.createTextNode (doc_, '');
   nd_txt_       Dbms_XmlDom.DomNode           := Dbms_XmlDom.makeNode (txt_el_);
   stream_       sys.utl_characterOutputStream := Dbms_XmlDom.setNodeValueAsCharacterStream (nd_txt_);
   throw_nd_     dbms_XmlDom.DomNode;
BEGIN
   WHILE offset_ <= Dbms_Lob.GetLength(blob_content_) LOOP
      Dbms_Lob.Read (blob_content_, chunk_size_, offset_, blob_chunk_);
      chunk_base64_ := Utl_Raw.Cast_To_Varchar2 (Utl_Encode.Base64_Encode (blob_chunk_));
      chunk_base64_ := replace (chunk_base64_, chr(13)||chr(10));
      chunk64_len_  := length (chunk_base64_);
      stream_.write (chunk_base64_, chunk64_len_);
      offset_ := offset_ + chunk_size_;
   END LOOP;
   stream_.close;
   throw_nd_ := Dbms_XmlDom.appendChild (
      Dbms_XmlDom.appendChild (append_to_, Make_Node(doc_,tag_name_,ns_,attrs_)),
      nd_txt_
   );
END Xml_Blob_Node;


------------
-- Silly Foundation stuff
PROCEDURE Init IS
BEGIN
   null;
END Init;

BEGIN
   Init;
END nyce_xml;
/
