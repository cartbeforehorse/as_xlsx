CREATE OR REPLACE PACKAGE nyce_utils IS

   ---
   -- Author  : Osian ap Garth
   -- Created : 2020-01-14
   -- Purpose :
   --    Facility-package to assist with boring foundational tasks that PL/SQL
   --    isn't so good at doing on its own
   -- 
   -- Developer Notes:
   -- Please observe coding conveitions:
   --   - 3-space indents
   --   - Uppercase oracle keywords
   --   - Camel-case underscored functions + procedures
   --   - Lowercase everything else
   --   - Trailing underscore for variables (none of this v_ or p_ silliness)
   --   - Stacked variables to appear indented beneath function calls, and not
   --     aligned with the end of a function; this avoids inconsistent indents
   --     and horizontal scrolling
   --   - Code with the next coder in mind.  Often, this will be you trying to
   --     debug your own mistakes, so be nice
   --

   DBMS_OUTPUT_SIZE_          CONSTANT NUMBER      := 1000000;


   ---------------------------
   -- Error handling functions
   --
   PROCEDURE Raise_App_Error (
      err_text_ IN VARCHAR2,
      p1_       IN VARCHAR2 := null,
      p2_       IN VARCHAR2 := null,
      p3_       IN VARCHAR2 := null,
      p4_       IN VARCHAR2 := null,
      p5_       IN VARCHAR2 := null,
      p6_       IN VARCHAR2 := null,
      p7_       IN VARCHAR2 := null,
      p8_       IN VARCHAR2 := null,
      p9_       IN VARCHAR2 := null,
      p0_       IN VARCHAR2 := null,
      repl_nl_  IN BOOLEAN  := true );

   ---------------------------
   -- General Oracle and PL/SQL helpers
   --
   FUNCTION Null_Date_Min (
      dt_ IN DATE ) RETURN DATE;
   FUNCTION Null_Date_Max (
      dt_ IN DATE ) RETURN DATE;
   FUNCTION Null_Date_Today (
      dt_    IN DATE,
      trunc_ IN VARCHAR2 := 'FALSE' ) RETURN DATE;


   ---------------------------
   -- Logging functionality
   --
   FUNCTION Strlim (
      str_      IN VARCHAR2,
      limit_    IN NUMBER := 4000 ) RETURN VARCHAR2;
   FUNCTION Trunc_Ellipsis (
      trunc_str_   IN VARCHAR2,
      trunc_len_   IN NUMBER  := 4000 ) RETURN VARCHAR2;

   FUNCTION Rep (
      msg_text_ IN CLOB,
      p1_       IN VARCHAR2 := null,
      p2_       IN VARCHAR2 := null,
      p3_       IN VARCHAR2 := null,
      p4_       IN VARCHAR2 := null,
      p5_       IN VARCHAR2 := null,
      p6_       IN VARCHAR2 := null,
      p7_       IN VARCHAR2 := null,
      p8_       IN VARCHAR2 := null,
      p9_       IN VARCHAR2 := null,
      p0_       IN VARCHAR2 := null,
      repl_nl_  IN BOOLEAN  := true ) RETURN CLOB;
   FUNCTION Reps (
      msg_     IN CLOB,
      p1_      IN VARCHAR2 := null,
      p2_      IN VARCHAR2 := null,
      p3_      IN VARCHAR2 := null,
      p4_      IN VARCHAR2 := null,
      p5_      IN VARCHAR2 := null,
      p6_      IN VARCHAR2 := null,
      p7_      IN VARCHAR2 := null,
      p8_      IN VARCHAR2 := null,
      p9_      IN VARCHAR2 := null,
      p0_      IN VARCHAR2 := null,
      repl_nl_ IN BOOLEAN  := true ) RETURN VARCHAR2;

   -- background job loggers
   PROCEDURE Trace (
      msg_     IN CLOB,
      p1_      IN VARCHAR2 := null,
      p2_      IN VARCHAR2 := null,
      p3_      IN VARCHAR2 := null,
      p4_      IN VARCHAR2 := null,
      p5_      IN VARCHAR2 := null,
      p6_      IN VARCHAR2 := null,
      p7_      IN VARCHAR2 := null,
      p8_      IN VARCHAR2 := null,
      p9_      IN VARCHAR2 := null,
      p0_      IN VARCHAR2 := null,
      repl_nl_ IN BOOLEAN  := true,
      quiet_   IN BOOLEAN  := false );
   PROCEDURE TraceQ (
      msg_     IN CLOB,
      p1_      IN VARCHAR2 := null,
      p2_      IN VARCHAR2 := null,
      p3_      IN VARCHAR2 := null,
      p4_      IN VARCHAR2 := null,
      p5_      IN VARCHAR2 := null,
      p6_      IN VARCHAR2 := null,
      p7_      IN VARCHAR2 := null,
      p8_      IN VARCHAR2 := null,
      p9_      IN VARCHAR2 := null,
      p0_      IN VARCHAR2 := null,
      repl_nl_ IN BOOLEAN  := true );
   PROCEDURE Log_Progress (
      msg_     IN CLOB,
      p1_      IN VARCHAR2 := null,
      p2_      IN VARCHAR2 := null,
      p3_      IN VARCHAR2 := null,
      p4_      IN VARCHAR2 := null,
      p5_      IN VARCHAR2 := null,
      p6_      IN VARCHAR2 := null,
      p7_      IN VARCHAR2 := null,
      p8_      IN VARCHAR2 := null,
      p9_      IN VARCHAR2 := null,
      p0_      IN VARCHAR2 := null,
      repl_nl_ IN BOOLEAN  := true );
   PROCEDURE Clear_Progress;
   PROCEDURE Log_Info (
      msg_     IN CLOB,
      p1_      IN VARCHAR2 := null,
      p2_      IN VARCHAR2 := null,
      p3_      IN VARCHAR2 := null,
      p4_      IN VARCHAR2 := null,
      p5_      IN VARCHAR2 := null,
      p6_      IN VARCHAR2 := null,
      p7_      IN VARCHAR2 := null,
      p8_      IN VARCHAR2 := null,
      p9_      IN VARCHAR2 := null,
      p0_      IN VARCHAR2 := null,
      repl_nl_ IN BOOLEAN  := true );
   PROCEDURE Log_Progress_And_Info (
      msg_     IN CLOB,
      p1_      IN VARCHAR2 := null,
      p2_      IN VARCHAR2 := null,
      p3_      IN VARCHAR2 := null,
      p4_      IN VARCHAR2 := null,
      p5_      IN VARCHAR2 := null,
      p6_      IN VARCHAR2 := null,
      p7_      IN VARCHAR2 := null,
      p8_      IN VARCHAR2 := null,
      p9_      IN VARCHAR2 := null,
      p0_      IN VARCHAR2 := null,
      repl_nl_ IN BOOLEAN  := true );
   PROCEDURE Log_Error (
      msg_     IN CLOB,
      p1_      IN VARCHAR2 := null,
      p2_      IN VARCHAR2 := null,
      p3_      IN VARCHAR2 := null,
      p4_      IN VARCHAR2 := null,
      p5_      IN VARCHAR2 := null,
      p6_      IN VARCHAR2 := null,
      p7_      IN VARCHAR2 := null,
      p8_      IN VARCHAR2 := null,
      p9_      IN VARCHAR2 := null,
      p0_      IN VARCHAR2 := null,
      repl_nl_ IN BOOLEAN  := true );

   -----
   -- environment checker
   --
   FUNCTION Env RETURN VARCHAR2;
   FUNCTION Is_Prod RETURN BOOLEAN;


END nyce_utils;
/
CREATE OR REPLACE PACKAGE BODY nyce_utils IS


------------------------------------------------------------------------------
------------------------------------------------------------------------------
--
-- General Functions to help with everyday tasks often required in PL/SQL
--
--

--------------------------------------
-- Error Handling
--
PROCEDURE Raise_App_Error (
   err_text_ IN VARCHAR2,
   p1_       IN VARCHAR2 := null,
   p2_       IN VARCHAR2 := null,
   p3_       IN VARCHAR2 := null,
   p4_       IN VARCHAR2 := null,
   p5_       IN VARCHAR2 := null,
   p6_       IN VARCHAR2 := null,
   p7_       IN VARCHAR2 := null,
   p8_       IN VARCHAR2 := null,
   p9_       IN VARCHAR2 := null,
   p0_       IN VARCHAR2 := null,
   repl_nl_  IN BOOLEAN  := true )
IS BEGIN
   Raise_Application_Error (-20110, Rep (
      err_text_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_
   ));
END Raise_App_Error;

--------------------------------------
-- Null_Date_Min(), Null_Date_Max()
--   Shorthand for the MIN and MAX date getters in IFS
FUNCTION Null_Date_Min (
   dt_ IN DATE ) RETURN DATE
IS BEGIN
   RETURN nvl (dt_, to_date('01/01/0001','DD/MM/YYYY'));
END Null_Date_Min;

FUNCTION Null_Date_Max (
   dt_ IN DATE ) RETURN DATE
IS BEGIN
   RETURN nvl (dt_, to_date('31/12/9999','DD/MM/YYYY'));
END Null_Date_Max;

FUNCTION Null_Date_Today (
   dt_    IN DATE,
   trunc_ IN VARCHAR2 := 'FALSE' ) RETURN DATE
IS
   out_dt_  DATE := nvl (dt_, sysdate);
BEGIN
   RETURN CASE
      WHEN trunc_='FALSE' THEN out_dt_
      ELSE trunc (out_dt_)
   END;
END Null_Date_Today;


------------------------------------------------------------
------------------------------------------------------------
--
-- Functions to help with IFS background-job logging
--
--

FUNCTION Strlim (
   str_      IN VARCHAR2,
   limit_    IN NUMBER := 4000 ) RETURN VARCHAR2
IS BEGIN
   RETURN substr (str_, 1, limit_);
END Strlim;

FUNCTION Trunc_Ellipsis (
   trunc_str_   IN VARCHAR2,
   trunc_len_   IN NUMBER  := 4000 ) RETURN VARCHAR2
IS BEGIN
   RETURN CASE
      WHEN length(trunc_str_) > trunc_len_ THEN substr (trunc_str_, trunc_len_-3) || '...'
      ELSE trunc_str_
   END;
END Trunc_Ellipsis;


FUNCTION Rep (
   msg_text_ IN CLOB,
   p1_       IN VARCHAR2 := null,
   p2_       IN VARCHAR2 := null,
   p3_       IN VARCHAR2 := null,
   p4_       IN VARCHAR2 := null,
   p5_       IN VARCHAR2 := null,
   p6_       IN VARCHAR2 := null,
   p7_       IN VARCHAR2 := null,
   p8_       IN VARCHAR2 := null,
   p9_       IN VARCHAR2 := null,
   p0_       IN VARCHAR2 := null,
   repl_nl_  IN BOOLEAN  := true ) RETURN CLOB
IS
   rtn_text_ VARCHAR2(32000);
BEGIN
   rtn_text_ := msg_text_;
   rtn_text_ := replace (rtn_text_, ':P1', p1_);
   rtn_text_ := replace (rtn_text_, ':P2', p2_);
   rtn_text_ := replace (rtn_text_, ':P3', p3_);
   rtn_text_ := replace (rtn_text_, ':P4', p4_);
   rtn_text_ := replace (rtn_text_, ':P5', p5_);
   rtn_text_ := replace (rtn_text_, ':P6', p6_);
   rtn_text_ := replace (rtn_text_, ':P7', p7_);
   rtn_text_ := replace (rtn_text_, ':P8', p8_);
   rtn_text_ := replace (rtn_text_, ':P9', p9_);
   rtn_text_ := replace (rtn_text_, ':P0', p0_);
   IF repl_nl_ THEN
      rtn_text_ := replace (rtn_text_, '<nl/>', utl_tcp.crlf);
      rtn_text_ := replace (rtn_text_, '<dnl/>', utl_tcp.crlf||utl_tcp.crlf);
   END IF;
   RETURN rtn_text_;
END Rep;

FUNCTION Reps (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true ) RETURN VARCHAR2
IS BEGIN
   RETURN  Dbms_Lob.Substr (
      Rep (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_),
      32000, 1
   );
END Reps;

-- Only useful for debugging
PROCEDURE Trace (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true,
   quiet_   IN BOOLEAN  := false )
IS
   logmsg_ VARCHAR2(32000) := Rep(msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_);
BEGIN
   Dbms_Output.Put_Line (CASE WHEN not quiet_ THEN 'Trace: ' END || logmsg_);
END Trace;

PROCEDURE TraceQ (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true )
IS BEGIN
   Trace (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_, true);
END TraceQ;

-- Doesn't really do much.  Legacy function that may have a purpose later.
PROCEDURE Log_Progress (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true )
IS
   PRAGMA AUTONOMOUS_TRANSACTION;
   logmsg_ VARCHAR2(32000) := strlim(Rep(msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_));
BEGIN
   Dbms_Output.Put_Line('Progress: ' || logmsg_);
   commit;
END Log_Progress;

PROCEDURE Clear_Progress
IS BEGIN
   Log_Progress ('');
END Clear_Progress;


-- Info message in the details section of background jobs
PROCEDURE Log_Info (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true )
IS
   PRAGMA AUTONOMOUS_TRANSACTION;
   logmsg_ VARCHAR2(32000) := strlim(Rep(msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_));
BEGIN
   Dbms_Output.Put_Line ('Info: ' || logmsg_);
   commit;
END Log_Info;

PROCEDURE Log_Progress_And_Info (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_       IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true )
IS BEGIN
   Log_Progress (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_);
   Log_Info (msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_);
END Log_Progress_And_Info;

-- Warning message in the details section of background jobs
PROCEDURE Log_Error (
   msg_     IN CLOB,
   p1_      IN VARCHAR2 := null,
   p2_      IN VARCHAR2 := null,
   p3_      IN VARCHAR2 := null,
   p4_      IN VARCHAR2 := null,
   p5_      IN VARCHAR2 := null,
   p6_      IN VARCHAR2 := null,
   p7_      IN VARCHAR2 := null,
   p8_      IN VARCHAR2 := null,
   p9_      IN VARCHAR2 := null,
   p0_      IN VARCHAR2 := null,
   repl_nl_ IN BOOLEAN  := true )
IS
   PRAGMA AUTONOMOUS_TRANSACTION;
   logmsg_ VARCHAR2(32000) := strlim(Rep(msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_));
BEGIN
   --Transaction_SYS.Set_Status_Info (logmsg_, 'WARNING', write_key_value_ => true);
   Dbms_Output.Put_Line('Error: ' || logmsg_);
   COMMIT;
END Log_Error;


PROCEDURE Purge_Dbms_Output
IS
   PRAGMA AUTONOMOUS_TRANSACTION;
BEGIN
   /*
    * Following code loops around the DBMS buffer, should you want to
    * archive it off to a table or an external LOG file or something
      LOOP
         Dbms_Output.Get_Line (line_text_, status_);
         EXIT WHEN status_ = 1;
      END LOOP;   
    */
   Dbms_Output.Disable;
   Dbms_Output.Enable (DBMS_OUTPUT_SIZE_);
   COMMIT;
END Purge_Dbms_Output;
--
-- End Logging Functions
------------------------------------------------------------
------------------------------------------------------------

-----------------------------------------------------
-- Environment checker
--
FUNCTION Env RETURN VARCHAR2
IS
   env_  VARCHAR2(20);
BEGIN
   SELECT d.name INTO env_ FROM v$database d;
   RETURN env_;
END Env;

FUNCTION Is_Prod RETURN BOOLEAN
IS BEGIN
   RETURN Env = 'LIVE';
END Is_Prod;

END nyce_utils;
/
