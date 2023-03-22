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


namespace NPM_1.Controllers
{
    public class TXPOController : Controller
    {
        // GET: TXPO
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult TXPO_Measurements()
        {
            return View();
        }



        public Excel.Application xlApp { get; set; }
        public Excel.Workbook Source_workbook { get; set; }

        public string Import_S_First = "";
        public string Import_S_Second = "";

        [HttpPost]
        // [ActionName("Upload_BSC_Data1")]
        public JsonResult Index_Post()
        {
            bool flag = true;
            string responseMessage = string.Empty;

            if (Request.Files.Count > 0)
            {
                HttpPostedFileBase file = Request.Files[0];

                //add more conditions like file type, file size etc as per your need.
                if (file != null && file.ContentLength > 0 && (Path.GetExtension(file.FileName).ToLower() == ".xlsx" || Path.GetExtension(file.FileName).ToLower() == ".xls" || Path.GetExtension(file.FileName).ToLower() == ".csv"))
                {
                    try
                    {
                        string fileName = Path.GetFileName(file.FileName);
                        string filePath = Path.Combine(Server.MapPath("~/UploadFiles"), fileName);
                        file.SaveAs(filePath);

                        flag = true;
                        responseMessage = "Upload Successful.";


                        //var Source_workbook = new XLWorkbook(filePath, XLEventTracking.Disabled);
                        //IXLWorksheet Source_worksheet = null;
                        //string Table_Name = Source_workbook.Worksheet(1).Name.ToString();


                        xlApp = new Excel.Application();
                        Source_workbook = xlApp.Workbooks.Open(filePath);
                        Excel.Worksheet sheet1 = Source_workbook.Worksheets[1];

                        string Table_Name = sheet1.Name;
                        Source_workbook.Close();

                        string ConnectionString = @"Server=NAKPRG-NB1243\AHMAD; Database=TXPO; Trusted_Connection=True;";
                        SqlConnection connection = new SqlConnection(ConnectionString);
                        connection.Open();


                        if (Table_Name== "Ericsson_EBAND")
                        {

                            // Uplaod File into SQL
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[0%-5%],
[5%-10%],
[10%-15%],
[15%-20%],
[20%-25%],
[25%-30%],
[30%-35%],
[35%-40%],
[40%-45%],
[45%-50%],
[50%-55%],
[55%-60%],
[60%-65%],
[65%-70%],
[70%-75%],
[75%-80%],
[80%-85%],
[85%-90%],
[90%-95%],
[95%-100%]  INTO[") + Table_Name;
                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"] select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[0%-5%],
[5%-10%],
[10%-15%],
[15%-20%],
[20%-25%],
[25%-30%],
[30%-35%],
[35%-40%],
[40%-45%],
[45%-50%],
[50%-55%],
[55%-60%],
[60%-65%],
[65%-70%],
[70%-75%],
[75%-80%],
[80%-85%],
[85%-90%],
[90%-95%],
[95%-100%] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by EndTime", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }





                        if (Table_Name == "Ericsson_HRAN")
                        {

                            // Uplaod File into SQL
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[Utilization] INTO[") + Table_Name;
                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"] select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[Utilization] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by EndTime", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }





                        if (Table_Name == "Ericsson_LRAN")
                        {

                            // Uplaod File into SQL
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[Utilization] INTO[") + Table_Name;
                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"] select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[Utilization] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by EndTime", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }





                        if (Table_Name == "Ericsson_TN")
                        {

                            // Uplaod File into SQL
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[Avg],
[Max],
[Min],
[0%-5%],
[5%-10%],
[10%-15%],
[15%-20%],
[20%-25%],
[25%-30%],
[30%-35%],
[35%-40%],
[40%-45%],
[45%-50%],
[50%-55%],
[55%-60%],
[60%-65%],
[65%-70%],
[70%-75%],
[75%-80%],
[80%-85%],
[85%-90%],
[90%-95%],
[95%-100%] INTO[") + Table_Name;
                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"] select [NeId],
[NeAlias],
[NeType],
[EntityType],
[MeasurePoint],
[EndTime],
[Failure],
[Avg],
[Max],
[Min],
[0%-5%],
[5%-10%],
[10%-15%],
[15%-20%],
[20%-25%],
[25%-30%],
[30%-35%],
[35%-40%],
[40%-45%],
[45%-50%],
[50%-55%],
[55%-60%],
[60%-65%],
[65%-70%],
[70%-75%],
[75%-80%],
[80%-85%],
[85%-90%],
[90%-95%],
[95%-100%] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by EndTime", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }



                        if (Table_Name == "HRAN PORT BANDWIDTH UTIL Daily")
                        {
                            //string Table_Name1 = "Huawei_HRAN";
                            // Uplaod File into SQL
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [Date],
[ElementID],
[ElementID1],
[MAXIMUM PORT_RX_BW_UTILIZATION %],
[MAXIMUM PORT_TX_BW_UTILIZATION %],
[AVERAGE PORT_RX_BW_UTILIZATION %],
[AVERAGE PORT_TX_BW_UTILIZATION %],
[AVG TOP8 PORT_RX_BW_UTILIZATION],
[AVG TOP8 PORT_TX_BW_UTILIZATION] INTO['") + Table_Name;
                            string IMPORT_STR_2 = string.Format(@"$'] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO ['") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"$'] select [Date],
[ElementID],
[ElementID1],
[MAXIMUM PORT_RX_BW_UTILIZATION %],
[MAXIMUM PORT_TX_BW_UTILIZATION %],
[AVERAGE PORT_RX_BW_UTILIZATION %],
[AVERAGE PORT_TX_BW_UTILIZATION %],
[AVG TOP8 PORT_RX_BW_UTILIZATION],
[AVG TOP8 PORT_TX_BW_UTILIZATION] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }



                        if (Table_Name == "LRAN PORT BANDWIDTH UTILIZATION")
                        {
                            //string Table_Name1 = "Huawei_HRAN";
                            // Uplaod File into SQL
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [Date],
[ElementID],
[ElementID1],
[MAXIMUM PORT_RX_BW_UTILIZATION %],
[MAXIMUM PORT_TX_BW_UTILIZATION %],
[AVERAGE PORT_RX_BW_UTILIZATION %],
[AVERAGE PORT_TX_BW_UTILIZATION %],
[AVG TOP8 PORT_RX_BW_UTILIZATION],
[AVG TOP8 PORT_TX_BW_UTILIZATION] INTO['") + Table_Name;
                            string IMPORT_STR_2 = string.Format(@"$'] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO ['") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"$'] select [Date],
[ElementID],
[ElementID1],
[MAXIMUM PORT_RX_BW_UTILIZATION %],
[MAXIMUM PORT_TX_BW_UTILIZATION %],
[AVERAGE PORT_RX_BW_UTILIZATION %],
[AVERAGE PORT_TX_BW_UTILIZATION %],
[AVG TOP8 PORT_RX_BW_UTILIZATION],
[AVG TOP8 PORT_TX_BW_UTILIZATION] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
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
                            string tb = (Name_Table.Rows[k]).ItemArray[0].ToString();
                            if (tb.Substring(tb.Length-2,1)=="$")
                            {
                                tb = tb.Substring(1, tb.Length - 3);
                            }
                            if (tb == Table_Name)
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