CREATE OR REPLACE PACKAGE CBH_UTILS_API IS

   module_  CONSTANT VARCHAR2(25) := 'FNDBAS';
   lu_name_ CONSTANT VARCHAR2(25) := 'CbhUtils';
   lu_type_ CONSTANT VARCHAR2(25) := 'Custom';

   TYPE pipeline_rec IS RECORD (
      report_id       VARCHAR2(200),
      report_category VARCHAR2(200),
      db_name         VARCHAR2(20),
      user_id         VARCHAR2(30),
      db_timestamp    DATE,
      cli_timestamp   TIMESTAMP WITH TIME ZONE,
      chr_cli_tstamp  VARCHAR2(50),
      report_info     VARCHAR2(2000)
   );
   TYPE pipeline_table IS TABLE OF pipeline_rec;

   ---
   -- Author  : Osian ap Garth
   -- Created : 2020-01-14
   -- Purpose : Nice to have functions that act as a backbone to facilitate
   --           quicker development of often-used and general-purpose needs
   --           while developting IFS-compatible APIs
   -- 
   -- Developer Notes:
   -- Please observe coding conveitions:
   --   - 3-space indents
   --   - uppercase oracle keywords
   --   - camel-case underscored functions + procedures
   --   - lowercase everything else
   --   - trailing underscore for variables (none of this v_ or p_ silliness)
   --   - try to indent in a way that gives the next person a chance to understand your code!
   --

   DBMS_OUTPUT_SIZE_          CONSTANT NUMBER      := 1000000;

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


   -----------------------------------------------------
   -- IFS Foundation Stuff
   --
   PROCEDURE Init;


END CBH_UTILS_API;
/
CREATE OR REPLACE PACKAGE BODY CBH_UTILS_API IS


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

-- Because the IFS standard replace only allows 3 variables, which can be limiting
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
   logmsg_ VARCHAR2(32000) := strlim(Rep(msg_, p1_, p2_, p3_, p4_, p5_, p6_, p7_, p8_, p9_, p0_, repl_nl_));
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

-- Progress box on IFS background-job header
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
   COMMIT;
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
   COMMIT;
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


-----------------------------------------------------
-- IFS Foundation Stuff
--
PROCEDURE Init IS
BEGIN
   NULL;
END Init;

BEGIN
   Init;
END CBH_UTILS_API;
/
