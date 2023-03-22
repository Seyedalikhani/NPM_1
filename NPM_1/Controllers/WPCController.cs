using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using Newtonsoft.Json;
using System.Web.Services;
using System.Configuration;
using DocumentFormat.OpenXml.Spreadsheet;


namespace NPM_1.Controllers
{
    public class WPCController : Controller
    {
        // GET: WPC
        public ActionResult WPC_KPI()
        {
            return View();

        }

        public Excel.Application xlApp { get; set; }
        public Excel.Workbook Source_workbook { get; set; }
        public string Import_S_First = "";
        public string Import_S_Second = "";


        [HttpPost]
        public JsonResult Index_Post()
        {
            bool flag = true;
            string responseMessage = string.Empty;

            if (Request.Files.Count > 0)
            {
                HttpPostedFileBase file = Request.Files[0];

                //add more conditions like file type, file size etc as per your need.
                if (file != null && file.ContentLength > 0 && (Path.GetExtension(file.FileName).ToLower() == ".xlsb" || Path.GetExtension(file.FileName).ToLower() == ".xlsx" || Path.GetExtension(file.FileName).ToLower() == ".xls"))
                {
                    try
                    {
                        string fileName = Path.GetFileName(file.FileName);
                        string filePath = Path.Combine(Server.MapPath("~/UploadFiles"), fileName);
                        file.SaveAs(filePath);

                        flag = true;
                        responseMessage = "Upload Successful.";


                        xlApp = new Excel.Application();
                        Source_workbook = xlApp.Workbooks.Open(filePath);
                        int numSheets = Source_workbook.Worksheets.Count;
                        string Table_Name1 = "";
                        string Table_Name2 = "";
                        string Table_Name3 = "";

                        int y1 = 1;
                        Excel.Worksheet sheet1 = Source_workbook.Worksheets[y1];
                        while (sheet1.Name.ToString() != "CC2 Eric Cell BH")
                        {
                            y1++;
                            sheet1 = Source_workbook.Worksheets[y1];
                        }
                        Table_Name1 = sheet1.Name;


                        int y2 = 1;
                        Excel.Worksheet sheet2 = Source_workbook.Worksheets[y2];
                        while (sheet2.Name.ToString() != "CC2 Huawei Cell BH")
                        {
                            y2++;
                            sheet2 = Source_workbook.Worksheets[y2];
                        }
                        Table_Name2 = sheet2.Name;



                        int y3 = 1;
                        Excel.Worksheet sheet3 = Source_workbook.Worksheets[y3];
                        while (sheet3.Name.ToString() != "CC2 NOKIA SEG BH")
                        {
                            y3++;
                            sheet3 = Source_workbook.Worksheets[y3];
                        }
                        Table_Name3 = sheet3.Name;


                        int Ericsson_CC2 = 1; //1
                        int Huawei_CC2 = 0;  //2
                        int Nokia_CC2 = 0;  //3
                        int Ericsson_CC3 = 0;  //1
                        int Huawei_CC3 = 0;  //2
                        int Nokia_CC3 = 0;  //3
                        int Ericsson_RD3 = 0;  //1
                        int Huawei_RD3 = 0;  //2
                        int Nokia_RD3 = 0;  //3 
                        int Ericsson_RD4 = 0;  //1
                        int Huawei_RD4 = 0;  //3
                        int Nokia_RD4 = 0;  //2



                        string ConnectionString = @"Server=NAKPRG-NB1243\AHMAD; Database=WMT; Trusted_Connection=True;";
                        SqlConnection connection = new SqlConnection(ConnectionString);
                        connection.Open();



                        // 2G
                        if (Ericsson_CC2 == 1 && Table_Name1 == "CC2 Eric Cell BH")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243\AHMAD; Database=WMT; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [BSC], [CELL], [REGION], [PROVINCE], [Date], [CSSR_MCI], [OHSR], [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] INTO [CC2_TBL_Ericsson]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name1 + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [CC2_TBL_Ericsson]";
                            string IMPORT_STR_5 = string.Format(@" select [BSC], [CELL], [REGION], [PROVINCE], [Date], [CSSR_MCI], [OHSR], [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name1 + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }







                        // reding the Table name lists to determine what command must be run First or Second
                        string Table_LIST_STR = string.Format(@"select TABLE_NAME from INFORMATION_SCHEMA.TABLES");
                        SqlCommand Table_LIST_Command = new SqlCommand(Table_LIST_STR, connection);
                        Table_LIST_Command.ExecuteNonQuery();

                        DataTable Name_Table = new DataTable();
                        SqlDataAdapter dataAdapter = new SqlDataAdapter(Table_LIST_Command);
                        dataAdapter.Fill(Name_Table);

                        int table_finder = 0;
                        for (int k = 0; k < Name_Table.Rows.Count; k++)
                        {
                            if ((Name_Table.Rows[k]).ItemArray[0].ToString() == Table_Name1)
                            {
                                table_finder++;
                                SqlCommand Import_command = new SqlCommand(Import_S_Second, connection);
                                Import_command.ExecuteNonQuery();
                                break;
                            }
                            if ((Name_Table.Rows[k]).ItemArray[0].ToString() == "Medical_2G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Medical_4G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Medical_3G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "CC3_TBL_Ericsson" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "CC3_TBL_Huawei" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "CC3_TBL_Nokia" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "RD3_TBL_Ericsson" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "RD3_TBL_Huawei" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "RD3_TBL_Nokia")
                            {
                                table_finder++;
                                SqlCommand Import_command = new SqlCommand(Import_S_Second, connection);
                                Import_command.ExecuteNonQuery();
                                break;
                            }
                        }
                        if (table_finder == 0)
                        {
                            SqlCommand Import_command = new SqlCommand(Import_S_First, connection);
                            Import_command.ExecuteNonQuery();
                        }




                    }
                    catch (Exception ex)
                    {
                        flag = false;
                        responseMessage = "Upload Failed with error: " + ex.Message;
                    }
                }
                else
                {
                    flag = false;
                    responseMessage = "File is invalid.";
                }
            }
            else
            {
                flag = false;
                responseMessage = "File Upload has no file.";
            }

            return Json(new { success = flag, responseMessage = responseMessage }, JsonRequestBehavior.AllowGet);
        }








    }
}