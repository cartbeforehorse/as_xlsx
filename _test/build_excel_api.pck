CREATE OR REPLACE PACKAGE Build_Excel_API IS

   PROCEDURE Make;

END Build_Excel_API;
/
CREATE OR REPLACE PACKAGE BODY Build_Excel_API IS

PROCEDURE Make
IS
   report_params_ as_xlsx.params_arr := as_xlsx.params_arr();
   filename_      VARCHAR2(200) := Cbh_Utils_API.Rep ('TestExcel_:P1.xlsx', to_char(sysdate,'YYYYMMDD-HH24MI'));
   xl_            BLOB;
BEGIN

   As_Xlsx.Init_Workbook;

   report_params_.EXTEND(3);
   report_params_(1) := as_xlsx.param_rec ('Customer Id', '123', 'Bob the customer');
   report_params_(2) := as_xlsx.param_rec ('Date From',  '01-01-2024', '');
   report_params_(3) := as_xlsx.param_rec ('Date Until', '31-12-2024', '');

   ----
   -- Initiate the Excel sheet
   As_Xlsx.Set_Sheet_Name (1, 'Parameters');
   --Init_Fonts_And_Fills;
   --Create_Params_Sheet (customer_id_, date_from_, date_to_);
   As_Xlsx.Create_Params_Sheet ('My Report Parameters', report_params_);

   As_Xlsx.Comment (3, 3, 'This is a silly dingle dongle', 'Bob the Builder', 300, 200, 1);

   As_Xlsx.CellN (2, 12, 123.456, numFmtId_ => 'gbp_curr2');

   -- Save the file to disk and send by mail
   xl_ := As_Xlsx.Finish;
   Dbms_Output.Put_Line ('File created: ' || filename_);

   As_Xlsx.Save (xl_, 'EXCEL_OUT', filename_);
   Dbms_Output.Put_Line (filename_ || ' saved to filesystem');

END Make;

END Build_Excel_API;
/
