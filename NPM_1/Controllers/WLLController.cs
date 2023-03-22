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
    public class WLLController : Controller
    {
        // GET: WLL
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult WLL_KPI()
        {


            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            string Cell_Select = "select* from[dbo].[KPI_AG_Daily] where[UserLabel]='AG1G1700A' order by Date";
            SqlCommand Cell_Select1 = new SqlCommand(Cell_Select, connection);
            Cell_Select1.ExecuteNonQuery();

            DataTable Cell_Select_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(Cell_Select1);
            dataAdapter.Fill(Cell_Select_Table);

            List<DataPoint> dataPoints = new List<DataPoint>();

            for (int k = 0; k <= Cell_Select_Table.Rows.Count - 1; k++)
            {
                string date = Cell_Select_Table.Rows[k].ItemArray[2].ToString();
                //DateTime dt = Convert.ToDateTime(date);
                //double dt1 = dt.Ticks;
                string kpi = Cell_Select_Table.Rows[k].ItemArray[13].ToString();
                double kpi1 = Convert.ToDouble(kpi);
                // dataPoints.Add(new DataPoint(dt1, kpi1));

                TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
                dataPoints.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1));
            }

            ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoints);

            return View();
        }


        public class DataPoint
        {
            public DataPoint(double x, double y)
            {
                this.x = x;
                this.y = y;
            }

            //Explicitly setting the name to be used while serializing to JSON.
            //[DataMember(Name = "x")]
            public Nullable<double> x = null;

            //Explicitly setting the name to be used while serializing to JSON.
            //[DataMember(Name = "y")]
            public Nullable<double> y = null;
        }





        public ActionResult getProvince()
        {
            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            List<Province> PROVINCE_LIST = new List<Province>();
            string Province_Name = "";
            for (int r = 1; r <= 3; r++)
            {
                if (r == 1)
                {
                    Province_Name = "West Azarbaijan";
                }
                if (r == 2)
                {
                    Province_Name = "East Azarbaijan";
                }
                if (r == 3)
                {
                    Province_Name = "Kuzestan";
                }
                PROVINCE_LIST.Add(new Province
                {
                    ProvinceName = Province_Name
                });
            }

            return Json(PROVINCE_LIST, JsonRequestBehavior.AllowGet);
        }


        //public ActionResult getBSC()
        //{

        //    string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
        //    SqlConnection connection = new SqlConnection(ConnectionString);
        //    connection.Open();


        //    string BSC_STR_1 = "select distinct substring([Object identifier],1,5) from[KPI_AG_Daily]";
        //    SqlCommand BSC_Command1 = new SqlCommand(BSC_STR_1, connection);
        //    BSC_Command1.ExecuteNonQuery();

        //    DataTable BSC_Table = new DataTable();
        //    SqlDataAdapter dataAdapter = new SqlDataAdapter(BSC_Command1);
        //    dataAdapter.Fill(BSC_Table);

        //    List<BSC> BSC_LIST = new List<BSC>();
        //    string BSC_Name = "";
        //    for (int k = 1; k <= BSC_Table.Rows.Count; k++)
        //    {
        //        BSC_Name = (BSC_Table.Rows[k - 1]).ItemArray[0].ToString();
        //        BSC_LIST.Add(new BSC
        //        {
        //            BSCName = BSC_Name
        //        });
        //    }

        //    return Json(BSC_LIST, JsonRequestBehavior.AllowGet);
        //}


        //public ActionResult getCell()
        //{

        //    string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
        //    SqlConnection connection = new SqlConnection(ConnectionString);
        //    connection.Open();


        //    string Cell_STR_1 = "select distinct [UserLabel] from [KPI_AG_Daily]";
        //    SqlCommand Cell_Command1 = new SqlCommand(Cell_STR_1, connection);
        //    Cell_Command1.ExecuteNonQuery();

        //    DataTable Cell_Table = new DataTable();
        //    SqlDataAdapter dataAdapter = new SqlDataAdapter(Cell_Command1);
        //    dataAdapter.Fill(Cell_Table);

        //    List<Cell> Cell_LIST = new List<Cell>();
        //    string Cell_Name = "";
        //    for (int k = 1; k <= Cell_Table.Rows.Count; k++)
        //    {
        //        Cell_Name = (Cell_Table.Rows[k - 1]).ItemArray[0].ToString();
        //        Cell_LIST.Add(new Cell
        //        {
        //            CellName = Cell_Name
        //        });
        //    }

        //    return Json(Cell_LIST, JsonRequestBehavior.AllowGet);

        //}




        //public ActionResult getKPI()
        //{

        //    string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
        //    SqlConnection connection = new SqlConnection(ConnectionString);
        //    connection.Open();


        //    string KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'KPI_AG_Daily'";
        //    SqlCommand KPI_Command1 = new SqlCommand(KPI_STR_1, connection);
        //    KPI_Command1.ExecuteNonQuery();

        //    DataTable KPI_Table = new DataTable();
        //    SqlDataAdapter dataAdapter = new SqlDataAdapter(KPI_Command1);
        //    dataAdapter.Fill(KPI_Table);

        //    List<KPI> KPI_LIST = new List<KPI>();
        //    string KPI_Name = "";
        //    for (int k = 7; k <= KPI_Table.Rows.Count; k++)
        //    {
        //        KPI_Name = (KPI_Table.Rows[k - 1]).ItemArray[0].ToString();
        //        KPI_LIST.Add(new KPI
        //        {
        //            KPIName = KPI_Name
        //        });
        //    }

        //    return Json(KPI_LIST, JsonRequestBehavior.AllowGet);
        //}






        public Excel.Application xlApp { get; set; }
        public Excel.Workbook Source_workbook { get; set; }

        public string Import_S_First = "";
        public string Import_S_Second = "";
        private object dropdown_Data_type;


        public string Server_Name = @"NAKPRG-NB1243";
        public string DataBase_Name = "Data";

        //public string Server_Name = "172.26.7.159";
        //public string DataBase_Name = "Performance_NAK";


        [HttpPost]
        public JsonResult dropdown_provincePost_Node(string text)
        {
            bool flag = true;
            string responseMessage = string.Empty;
            string interval_province_name = text;
            //  return Json(new { success = true, result = text });

            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            string BSC_STR_1 = "";
            if (interval_province_name == "Daily/West Azarbaijan")
            {
                BSC_STR_1 = "select distinct substring([Object identifier],1,5) from [KPI_AG_Daily]";
            }
            if (interval_province_name == "BH/West Azarbaijan")
            {
                BSC_STR_1 = "select distinct substring([Object identifier],1,5) from [KPI_AG_BH]";
            }
            if (interval_province_name == "Daily/East Azarbaijan")
            {
                BSC_STR_1 = "select distinct [BSC] from [KPI_AS_Daily]";
            }
            if (interval_province_name == "BH/East Azarbaijan")
            {
                BSC_STR_1 = "select distinct [BSC] from [KPI_AS_BH]";
            }

            SqlCommand BSC_Command1 = new SqlCommand(BSC_STR_1, connection);
            BSC_Command1.ExecuteNonQuery();

            DataTable BSC_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(BSC_Command1);
            dataAdapter.Fill(BSC_Table);

            List<BSC> BSC_LIST = new List<BSC>();
            string BSC_Name = "";
            for (int k = 1; k <= BSC_Table.Rows.Count; k++)
            {
                BSC_Name = (BSC_Table.Rows[k - 1]).ItemArray[0].ToString();
                if (BSC_Name == "BSC411-Ahar(9)")
                {
                    BSC_Name = "B411W";
                }
                if (BSC_Name == "B413W-Maraghe(4)")
                {
                    BSC_Name = "B413W";
                }
                if (BSC_Name == "B414W_Marand(7)")
                {
                    BSC_Name = "B414W";
                }
                if (BSC_Name == "BSC418(Hashtroud+Miyanhe)(5)")
                {
                    BSC_Name = "B418W";
                }
                if (BSC_Name == "B417W(Tabriz-Kaleybar)(1)")
                {
                    BSC_Name = "";
                }
                if (BSC_Name != "" && BSC_Name != null)
                {
                    BSC_LIST.Add(new BSC
                    {
                        BSCName = BSC_Name
                    });
                }
            }

            return Json(BSC_LIST, JsonRequestBehavior.AllowGet);

        }


        [HttpPost]
        public JsonResult dropdown_provincePost_Cell(string text)
        {
            bool flag = true;
            string responseMessage = string.Empty;
            string interval_province_name = text;
            //  return Json(new { success = true, result = text });

            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            string Cell_STR_1 = "";
            if (interval_province_name == "Daily/West Azarbaijan")
            {
                Cell_STR_1 = "select distinct [UserLabel] from [KPI_AG_Daily]";
            }
            if (interval_province_name == "BH/West Azarbaijan")
            {
                Cell_STR_1 = "select distinct [UserLabel] from [KPI_AG_BH]";
            }
            if (interval_province_name == "Daily/East Azarbaijan")
            {
                Cell_STR_1 = "select distinct [Bts] from [KPI_AS_Daily]";
            }
            if (interval_province_name == "BH/East Azarbaijan")
            {
                Cell_STR_1 = "select distinct [Bts] from [KPI_AS_BH]";
            }

            SqlCommand Cell_Command1 = new SqlCommand(Cell_STR_1, connection);
            Cell_Command1.ExecuteNonQuery();

            DataTable Cell_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(Cell_Command1);
            dataAdapter.Fill(Cell_Table);

            List<Cell> Cell_LIST = new List<Cell>();
            string Cell_Name = "";
            for (int k = 1; k <= Cell_Table.Rows.Count; k++)
            {
                Cell_Name = (Cell_Table.Rows[k - 1]).ItemArray[0].ToString();
                Cell_LIST.Add(new Cell
                {
                    CellName = Cell_Name
                });
            }

            return Json(Cell_LIST, JsonRequestBehavior.AllowGet);
        }




        [HttpPost]
        public JsonResult dropdown_provincePost_KPI(string text)
        {
            bool flag = true;
            string responseMessage = string.Empty;
            string interval_province_name = text;
            //  return Json(new { success = true, result = text });

            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            string KPI_STR_1 = "";
            int KPI_Start_Index = 0;
            if (interval_province_name == "Daily/West Azarbaijan")
            {
                KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'KPI_AG_Daily'";
                KPI_Start_Index = 7;
            }
            if (interval_province_name == "BH/West Azarbaijan")
            {
                KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'KPI_AG_BH'";
                KPI_Start_Index = 7;
            }
            if (interval_province_name == "Daily/East Azarbaijan")
            {
                KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'KPI_AS_Daily'";
                KPI_Start_Index = 6;
            }
            if (interval_province_name == "BH/East Azarbaijan")
            {
                KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'KPI_AS_BH'";
                KPI_Start_Index = 6;
            }

            SqlCommand KPI_Command1 = new SqlCommand(KPI_STR_1, connection);
            KPI_Command1.ExecuteNonQuery();

            DataTable KPI_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(KPI_Command1);
            dataAdapter.Fill(KPI_Table);

            List<KPI> KPI_LIST = new List<KPI>();
            string KPI_Name = "";
            for (int k = KPI_Start_Index; k <= KPI_Table.Rows.Count; k++)
            {
                KPI_Name = (KPI_Table.Rows[k - 1]).ItemArray[0].ToString();
                KPI_LIST.Add(new KPI
                {
                    KPIName = KPI_Name
                });
            }

            return Json(KPI_LIST, JsonRequestBehavior.AllowGet);
        }


        // Province Selection JQuary
        [HttpPost]
        public JsonResult dropdown_provincePost(string text)
        {
            bool flag = true;
            string responseMessage = string.Empty;
            string interval_province_kpi = text;

            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            int slash_ind = 0;
            int dash_ind = 0;
            for (int k = 0; k <= interval_province_kpi.Length - 1; k++)
            {
                if (interval_province_kpi[k].ToString() == "/")
                {
                    slash_ind = k;
                }
                if (interval_province_kpi[k].ToString() == "|")
                {
                    dash_ind = k;
                }
            }
            //string cell_name = interval_province_kpi.Substring(slash_ind + 1, interval_province_kpi.Length - slash_ind - 1);
            string kpi_name = "";
            if (interval_province_kpi[interval_province_kpi.Length - 1].ToString() != "|")
            {
                kpi_name = interval_province_kpi.Substring(dash_ind + 1, interval_province_kpi.Length - dash_ind - 1);
            }

            string Province_Select_Quary = "";
            string Table_Name = "";

            if (interval_province_kpi.Substring(0, dash_ind) == "Daily/West Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AG_PROVINCE_Daily] order by Date";
                Table_Name = "KPI_AG_PROVINCE_Daily";
            }
            if (interval_province_kpi.Substring(0, dash_ind) == "BH/West Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AG_PROVINCE_BH] order by Date";
                Table_Name = "KPI_AG_PROVINCE_BH";
            }
            if (interval_province_kpi.Substring(0, dash_ind) == "Daily/East Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AS_PROVINCE_Daily] order by [Begin time]";
                Table_Name = "KPI_AS_PROVINCE_Daily";
            }
            if (interval_province_kpi.Substring(0, dash_ind) == "BH/East Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AS_PROVINCE_BH] order by [Begin time]";
                Table_Name = "KPI_AS_PROVINCE_BH";
            }

            string Province_Select = Province_Select_Quary;
            SqlCommand Province_Select1 = new SqlCommand(Province_Select, connection);
            Province_Select1.ExecuteNonQuery();

            DataTable Province_Select_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(Province_Select1);
            dataAdapter.Fill(Province_Select_Table);

            List<DataPoint> dataPoints = new List<DataPoint>();

            // Find Index of KPI
            string KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " + "'" + Table_Name + "'";
            SqlCommand KPI_Command1 = new SqlCommand(KPI_STR_1, connection);
            KPI_Command1.ExecuteNonQuery();

            DataTable KPI_Table = new DataTable();
            SqlDataAdapter dataAdapter2 = new SqlDataAdapter(KPI_Command1);
            dataAdapter2.Fill(KPI_Table);

            int KPI_Ind = 0;
            for (int k = 1; k <= KPI_Table.Rows.Count; k++)
            {
                string KPI_Name = (KPI_Table.Rows[k - 1]).ItemArray[0].ToString();
                if (KPI_Name == kpi_name)
                {
                    KPI_Ind = k - 1;
                    break;
                }
            }
            if (kpi_name == "" || kpi_name == "Please Select a KPI")
            {
                KPI_Ind = 2;
            }

            if (KPI_Ind != 0)
            {
                for (int k = 0; k <= Province_Select_Table.Rows.Count - 1; k++)
                {
                    string date = Province_Select_Table.Rows[k].ItemArray[0].ToString();
                    //DateTime dt = Convert.ToDateTime(date);
                    //double dt1 = dt.Ticks;
                    string kpi = Province_Select_Table.Rows[k].ItemArray[KPI_Ind].ToString();
                    if (kpi != "")
                    {
                        double kpi1 = Convert.ToDouble(kpi);
                        // dataPoints.Add(new DataPoint(dt1, kpi1));

                        TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
                        dataPoints.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1));
                    }

                }

                //return Json(new { success = true, result = dataPoints });
            }
            return Json(new { success = true, result = dataPoints });
            //for (int k = 0; k <= Province_Select_Table.Rows.Count - 1; k++)
            //{
            //    string date = Province_Select_Table.Rows[k].ItemArray[0].ToString();
            //    //DateTime dt = Convert.ToDateTime(date);
            //    //double dt1 = dt.Ticks;
            //    string kpi = Province_Select_Table.Rows[k].ItemArray[KPI_Ind].ToString();
            //    if (kpi != "")
            //    {
            //        double kpi1 = Convert.ToDouble(kpi);
            //        // dataPoints.Add(new DataPoint(dt1, kpi1));

            //        TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
            //        dataPoints.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1));
            //    }

            //}

            //return Json(new { success = true, result = dataPoints });

        }





        // Node Selection JQuary





        // KPI Selection JQuary
        [HttpPost]
        public JsonResult dropdown_kpiPost(string text)
        {
            bool flag = true;
            string responseMessage = string.Empty;
            string interval_province_kpi = text;

            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            int slash_ind = 0;
            int dash_ind = 0;
            for (int k = 0; k <= interval_province_kpi.Length - 1; k++)
            {
                if (interval_province_kpi[k].ToString() == "/")
                {
                    slash_ind = k;
                }
                if (interval_province_kpi[k].ToString() == "|")
                {
                    dash_ind = k;
                }
            }
            //string cell_name = interval_province_kpi.Substring(slash_ind + 1, interval_province_kpi.Length - slash_ind - 1);
            string kpi_name = "";
            if (interval_province_kpi[interval_province_kpi.Length - 1].ToString() != "|")
            {
                kpi_name = interval_province_kpi.Substring(dash_ind + 1, interval_province_kpi.Length - dash_ind - 1);
            }

            string Province_Select_Quary = "";
            string Table_Name = "";

            if (interval_province_kpi.Substring(0, dash_ind) == "Daily/West Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AG_PROVINCE_Daily] order by Date";
                Table_Name = "KPI_AG_PROVINCE_Daily";
            }
            if (interval_province_kpi.Substring(0, dash_ind) == "BH/West Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AG_PROVINCE_BH] order by Date";
                Table_Name = "KPI_AG_PROVINCE_BH";
            }
            if (interval_province_kpi.Substring(0, dash_ind) == "Daily/East Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AS_PROVINCE_Daily] order by [Begin time]";
                Table_Name = "KPI_AS_PROVINCE_Daily";
            }
            if (interval_province_kpi.Substring(0, dash_ind) == "BH/East Azarbaijan")
            {
                Province_Select_Quary = "select * from [dbo].[KPI_AS_PROVINCE_BH] order by [Begin time]";
                Table_Name = "KPI_AS_PROVINCE_BH";
            }

            string Province_Select = Province_Select_Quary;
            SqlCommand Province_Select1 = new SqlCommand(Province_Select, connection);
            Province_Select1.ExecuteNonQuery();

            DataTable Province_Select_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(Province_Select1);
            dataAdapter.Fill(Province_Select_Table);

            List<DataPoint> dataPoints = new List<DataPoint>();

            // Find Index of KPI
            string KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " + "'" + Table_Name + "'";
            SqlCommand KPI_Command1 = new SqlCommand(KPI_STR_1, connection);
            KPI_Command1.ExecuteNonQuery();

            DataTable KPI_Table = new DataTable();
            SqlDataAdapter dataAdapter2 = new SqlDataAdapter(KPI_Command1);
            dataAdapter2.Fill(KPI_Table);

            int KPI_Ind = 0;
            for (int k = 1; k <= KPI_Table.Rows.Count; k++)
            {
                string KPI_Name = (KPI_Table.Rows[k - 1]).ItemArray[0].ToString();
                if (KPI_Name == kpi_name)
                {
                    KPI_Ind = k - 1;
                    break;
                }
            }
            if (kpi_name == "" || kpi_name == "Please Select a KPI")
            {
                KPI_Ind = 2;
            }

            for (int k = 0; k <= Province_Select_Table.Rows.Count - 1; k++)
            {
                string date = Province_Select_Table.Rows[k].ItemArray[0].ToString();
                //DateTime dt = Convert.ToDateTime(date);
                //double dt1 = dt.Ticks;
                string kpi = Province_Select_Table.Rows[k].ItemArray[KPI_Ind].ToString();
                if (kpi != "")
                {
                    double kpi1 = Convert.ToDouble(kpi);
                    // dataPoints.Add(new DataPoint(dt1, kpi1));

                    TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
                    dataPoints.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1));
                }

            }

            return Json(new { success = true, result = dataPoints });

        }






        // Cell Selection JQuary
        [HttpPost]
        public JsonResult dropdown_cellPost(string text)
        {
            bool flag = true;
            string responseMessage = string.Empty;
            string interval_province_kpi_cell_name = text;

            string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            int slash_ind = 0;
            int dash_ind = 0;
            for (int k = 0; k <= interval_province_kpi_cell_name.Length - 1; k++)
            {
                if (interval_province_kpi_cell_name[k].ToString() == "/")
                {
                    slash_ind = k;
                }
                if (interval_province_kpi_cell_name[k].ToString() == "|")
                {
                    dash_ind = k;
                }
            }
            string cell_name = interval_province_kpi_cell_name.Substring(slash_ind + 1, interval_province_kpi_cell_name.Length - slash_ind - 1);
            string kpi_name = interval_province_kpi_cell_name.Substring(dash_ind + 1, slash_ind - dash_ind - 1);
            string Cell_Select_Quary = "";
            string Table_Name = "";

            if (interval_province_kpi_cell_name.Substring(0, dash_ind) == "Daily/West Azarbaijan")
            {
                Cell_Select_Quary = "select * from [dbo].[KPI_AG_Daily] where [UserLabel]='" + cell_name + "' order by Date";
                Table_Name = "KPI_AG_Daily";
            }
            if (interval_province_kpi_cell_name.Substring(0, dash_ind) == "BH/West Azarbaijan")
            {
                Cell_Select_Quary = "select * from [dbo].[KPI_AG_BH] where [UserLabel]='" + cell_name + "' order by Date";
                Table_Name = "KPI_AG_BH";
            }
            if (interval_province_kpi_cell_name.Substring(0, dash_ind) == "Daily/East Azarbaijan")
            {
                Cell_Select_Quary = "select * from [dbo].[KPI_AS_Daily] where [Bts]='" + cell_name + "' order by [Begin time]";
                Table_Name = "KPI_AS_Daily";
            }
            if (interval_province_kpi_cell_name.Substring(0, dash_ind) == "BH/East Azarbaijan")
            {
                Cell_Select_Quary = "select * from [dbo].[KPI_AS_BH] where [Bts]='" + cell_name + "' order by [Begin time]";
                Table_Name = "KPI_AS_BH";
            }

            string Cell_Select = Cell_Select_Quary;
            SqlCommand Cell_Select1 = new SqlCommand(Cell_Select, connection);
            Cell_Select1.ExecuteNonQuery();

            DataTable Cell_Select_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(Cell_Select1);
            dataAdapter.Fill(Cell_Select_Table);

            List<DataPoint> dataPoints = new List<DataPoint>();

            // Find Index of KPI
            string KPI_STR_1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " + "'" + Table_Name + "'";
            SqlCommand KPI_Command1 = new SqlCommand(KPI_STR_1, connection);
            KPI_Command1.ExecuteNonQuery();

            DataTable KPI_Table = new DataTable();
            SqlDataAdapter dataAdapter2 = new SqlDataAdapter(KPI_Command1);
            dataAdapter2.Fill(KPI_Table);

            int KPI_Ind = 0;
            for (int k = 1; k <= KPI_Table.Rows.Count; k++)
            {
                string KPI_Name = (KPI_Table.Rows[k - 1]).ItemArray[0].ToString();
                if (KPI_Name == kpi_name)
                {
                    KPI_Ind = k - 1;
                    break;
                }
            }
            if (kpi_name == "Please Select a KPI" && Table_Name.Substring(4, 2) == "AG")
            {
                KPI_Ind = 6;
            }
            if (kpi_name == "Please Selece a KPI" && Table_Name.Substring(4, 2) == "AS")
            {
                KPI_Ind = 5;
            }

            for (int k = 0; k <= Cell_Select_Table.Rows.Count - 1; k++)
            {
                string date = Cell_Select_Table.Rows[k].ItemArray[2].ToString();
                //DateTime dt = Convert.ToDateTime(date);
                //double dt1 = dt.Ticks;
                string kpi = Cell_Select_Table.Rows[k].ItemArray[KPI_Ind].ToString();
                if (kpi != "")
                {
                    double kpi1 = Convert.ToDouble(kpi);
                    // dataPoints.Add(new DataPoint(dt1, kpi1));

                    TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
                    dataPoints.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1));
                }

            }

            return Json(new { success = true, result = dataPoints });

        }





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


                        //var Source_workbook = new XLWorkbook(filePath, XLEventTracking.Disabled);
                        //IXLWorksheet Source_worksheet = null;
                        //string Table_Name = Source_workbook.Worksheet(1).Name.ToString();



                        xlApp = new Excel.Application();
                        Source_workbook = xlApp.Workbooks.Open(filePath);
                        int numSheets = Source_workbook.Worksheets.Count;
                        string Table_Name = "";



                        Excel.Worksheet sheet1 = Source_workbook.Worksheets[1];
                        Table_Name = sheet1.Name;

                        //Excel.Worksheet sheet2 = Source_workbook.Worksheets[2];
                        //Table_Name = sheet2.Name;

                        //if (numSheets > 2)
                        //{
                        //    Excel.Worksheet sheet3 = Source_workbook.Worksheets[3];
                        //    // Table of Province
                        //    Table_Name = sheet3.Name;
                        //}


                        Source_workbook.Close();

                        string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
                        SqlConnection connection = new SqlConnection(ConnectionString);
                        connection.Open();




                        if (Table_Name == "KPI_AG_Daily")
                        {
                            // Uplaod File into SQL
                            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
                            //SqlConnection connection = new SqlConnection(ConnectionString);
                            //connection.Open();
                            //First Order to Make Table

                            string IMPORT_STR_1 = string.Format(@"select [Province ],
[City],
[Date],
[UserLabel],
[Object identifier],
[Cell and Location Area Cell(LAC-CI)],
[TCH in service rate(%)],
[TCH total traffic number],
[Carrier frequence number],
[Half Rate Usage(%)],
[TCH in congestion rate(include handover) (%)],
[TCH assign failure rate(%)],
[TCH total number in busy time],
[SDCCH total traffic number],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[SDCCH ASS SUCCESS RATE(%)],
[TCH in call drop rate(include handover) (%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)]
 INTO[") + Table_Name;

                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;

                            string IMPORT_STR_5 = string.Format(@"] select [Province ],
[City],
[Date],
[UserLabel],
[Object identifier],
[Cell and Location Area Cell(LAC-CI)],
[TCH in service rate(%)],
[TCH total traffic number],
[Carrier frequence number],
[Half Rate Usage(%)],
[TCH in congestion rate(include handover) (%)],
[TCH assign failure rate(%)],
[TCH total number in busy time],
[SDCCH total traffic number],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[SDCCH ASS SUCCESS RATE(%)],
[TCH in call drop rate(include handover) (%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)]
 from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }






                        if (Table_Name == "KPI_AG_Daily_Norooz98")
                        {
                            // Uplaod File into SQL
                            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
                            //SqlConnection connection = new SqlConnection(ConnectionString);
                            //connection.Open();
                            //First Order to Make Table

                            string IMPORT_STR_1 = string.Format(@"select [City],
[Date],
[UserLabel],
[Object identifier],
[Cell and Location Area Cell(LAC-CI)],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[TCH in service rate(%)],
[TCH in congestion rate(exclude handover)(%)],
[TCH in congestion rate(include handover)(%)],
[TCH in call drop rate(exclude handover)(%)],
[TCH in call drop rate(include handover)(%)],
[Handover success rate(%)],
[Dual band handover success rate(%)],
[SDCCH available total number],
[SDCCH attemption total number],
[SDCCH overflow total number],
[SDCCH Dropped call total number],
[SDCCH total traffic number],
[Available TCH total number],
[TCH attempt total number(exclude handover)],
[TCH overflow total number(exclude handover)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(exclude handover)],
[TCH seizure total number(include handover)],
[TCH dropped call total number],
[TCH total traffic number],
[Total number of handover request],
[Total number of successful handover],
[Total number of dual band handover attemption],
[Total number of successful dual band handover],
[TCH total traffic number of GSM1800],
[Busy cell number],
[Idle cell number],
[Bad cell number],
[The percent of bad cell],
[Frequence band],
[Traffic number of 24 hours],
[Traffic number per channel],
[SDCCH total number in busy time],
[TCH total number in busy time],
[Carrier frequence number],
[Whole carrier total number],
[radio switch rate(%)],
[Handover request rate in busy time(%)],
[TCH assign failure rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[Call success rate(%)],
[Call drop rate with handover],
[Call drop rate of traffic],
[holding time of averge call],
[SDCCH allocate failure number in busy time],
[TCH allocate failure number without handover],
[TCH allocate no success number without handover],
[TCH allocate failure number with handover],
[TCH allocate no success number with handover],
[TCH seize number in busy time],
[TCH allocate failure rate(%)],
[the traffic num from 0 to 1],
[the traffic num from 1 to 2],
[the traffic num from 2 to 3],
[the traffic num from 3 to 4],
[the traffic num from 4 to 5],
[the traffic num from 5 to 6],
[the traffic num from 6 to 7],
[the traffic num from 7 to 8],
[the traffic num from 8 to 9],
[the traffic num from 9 to 10],
[the traffic num from 10 to 11],
[the traffic num from 11 to 12],
[the traffic num from 12 to 13],
[the traffic num from 13 to 14],
[the traffic num from 14 to 15],
[the traffic num from 15 to 16],
[the traffic num from 16 to 17],
[the traffic num from 17 to 18],
[the traffic num from 18 to 19],
[the traffic num from 19 to 20],
[the traffic num from 20 to 21],
[the traffic num from 21 to 22],
[the traffic num from 22 to 23],
[the traffic num from 23 to 24]
 INTO[") + Table_Name;

                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;

                            string IMPORT_STR_5 = string.Format(@"] select [City],
[Date],
[UserLabel],
[Object identifier],
[Cell and Location Area Cell(LAC-CI)],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[TCH in service rate(%)],
[TCH in congestion rate(exclude handover)(%)],
[TCH in congestion rate(include handover)(%)],
[TCH in call drop rate(exclude handover)(%)],
[TCH in call drop rate(include handover)(%)],
[Handover success rate(%)],
[Dual band handover success rate(%)],
[SDCCH available total number],
[SDCCH attemption total number],
[SDCCH overflow total number],
[SDCCH Dropped call total number],
[SDCCH total traffic number],
[Available TCH total number],
[TCH attempt total number(exclude handover)],
[TCH overflow total number(exclude handover)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(exclude handover)],
[TCH seizure total number(include handover)],
[TCH dropped call total number],
[TCH total traffic number],
[Total number of handover request],
[Total number of successful handover],
[Total number of dual band handover attemption],
[Total number of successful dual band handover],
[TCH total traffic number of GSM1800],
[Busy cell number],
[Idle cell number],
[Bad cell number],
[The percent of bad cell],
[Frequence band],
[Traffic number of 24 hours],
[Traffic number per channel],
[SDCCH total number in busy time],
[TCH total number in busy time],
[Carrier frequence number],
[Whole carrier total number],
[radio switch rate(%)],
[Handover request rate in busy time(%)],
[TCH assign failure rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[Call success rate(%)],
[Call drop rate with handover],
[Call drop rate of traffic],
[holding time of averge call],
[SDCCH allocate failure number in busy time],
[TCH allocate failure number without handover],
[TCH allocate no success number without handover],
[TCH allocate failure number with handover],
[TCH allocate no success number with handover],
[TCH seize number in busy time],
[TCH allocate failure rate(%)],
[the traffic num from 0 to 1],
[the traffic num from 1 to 2],
[the traffic num from 2 to 3],
[the traffic num from 3 to 4],
[the traffic num from 4 to 5],
[the traffic num from 5 to 6],
[the traffic num from 6 to 7],
[the traffic num from 7 to 8],
[the traffic num from 8 to 9],
[the traffic num from 9 to 10],
[the traffic num from 10 to 11],
[the traffic num from 11 to 12],
[the traffic num from 12 to 13],
[the traffic num from 13 to 14],
[the traffic num from 14 to 15],
[the traffic num from 15 to 16],
[the traffic num from 16 to 17],
[the traffic num from 17 to 18],
[the traffic num from 18 to 19],
[the traffic num from 19 to 20],
[the traffic num from 20 to 21],
[the traffic num from 21 to 22],
[the traffic num from 22 to 23],
[the traffic num from 23 to 24]
 from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }

                                 



                        if (Table_Name == "KPI_AG_BSC_Daily")
                        {
                            // Uplaod File into SQL
                            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
                            //SqlConnection connection = new SqlConnection(ConnectionString);
                            //connection.Open();
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [DATE],
[BSC],
[City],
[TCH in service rate(%)],
[TCH total traffic number],
[Carrier frequence number],
[Half Rate Usage(%)],
[TCH in congestion rate(include handover) (%)],
[TCH assign failure rate(%)],
[TCH total number in busy time],
[SDCCH total traffic number],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[SDCCH ASS SUCCESS RATE(%)],
[TCH in call drop rate(include handover) (%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)]
 INTO[") + Table_Name;

                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"] select [DATE],
[BSC],
[City],
[TCH in service rate(%)],
[TCH total traffic number],
[Carrier frequence number],
[Half Rate Usage(%)],
[TCH in congestion rate(include handover) (%)],
[TCH assign failure rate(%)],
[TCH total number in busy time],
[SDCCH total traffic number],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[SDCCH ASS SUCCESS RATE(%)],
[TCH in call drop rate(include handover) (%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)]
 from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }






                        if (Table_Name == "KPI_AG_Province_Daily")
                        {
                            // Uplaod File into SQL
                            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
                            //SqlConnection connection = new SqlConnection(ConnectionString);
                            //connection.Open();
                            //First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [DATE],
[Province],
[TCH in service rate(%)],
[TCH total traffic number],
[Carrier frequence number],
[Half Rate Usage(%)],
[TCH in congestion rate(include handover) (%)],
[TCH assign failure rate(%)],
[TCH total number in busy time],
[SDCCH total traffic number],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[SDCCH ASS SUCCESS RATE(%)],
[TCH in call drop rate(include handover) (%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)]
 INTO[") + Table_Name;

                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"] select [DATE],
[Province],
[TCH in service rate(%)],
[TCH total traffic number],
[Carrier frequence number],
[Half Rate Usage(%)],
[TCH in congestion rate(include handover) (%)],
[TCH assign failure rate(%)],
[TCH total number in busy time],
[SDCCH total traffic number],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH channel in call drop rate(%)],
[SDCCH ASS SUCCESS RATE(%)],
[TCH in call drop rate(include handover) (%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)]
 from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;

                        }







                        if (Table_Name == "KPI_AS_Daily" || Table_Name == "KPI_AS_Hourly" || Table_Name == "KPI_AS_BH")
                        {
                            // Uplaod File into SQL
                            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
                            //SqlConnection connection = new SqlConnection(ConnectionString);
                            //connection.Open();
                            // First Order to Make Table
                            string IMPORT_STR_1 = string.Format(@"select [Province],
[City],
[Begin time],
[Bts],
[BSC],
[TCH in service rate(%)],
[TCH total traffic number (erl)],
[Whole carrier total number],
[Half rate Usage%],
[TCH in congestion rate(include handover)(%)],
[TCH assign failure rate(%)],
[Available TCH total number],
[SDCCH total traffic number (erl)],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH DROP Rate %],
[SDCCH assignments success rate%],
[TCH in call drop rate(include handover)(%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)],
[Call Set-up Success Rate(CSSR)(%)],
[Rx_Quality DL (0-5)],
[Rx_Quality UL (0-5)],
[TA=0 %],
[0<TA<=2 %],
[2<TA<=4 %],
[4<TA<=8 %],
[8<TA<=20 %],
[20<TA<=63 %],
[TA>63 %]
INTO[") + Table_Name;
                            string IMPORT_STR_2 = string.Format(@"] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$]", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;


                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = string.Format(@"INSERT INTO [") + Table_Name;
                            string IMPORT_STR_5 = string.Format(@"] select [Province],
[City],
[Begin time],
[Bts],
[BSC],
[TCH in service rate(%)],
[TCH total traffic number (erl)],
[Whole carrier total number],
[Half rate Usage%],
[TCH in congestion rate(include handover)(%)],
[TCH assign failure rate(%)],
[Available TCH total number],
[SDCCH total traffic number (erl)],
[SDCCH in service rate(%)],
[SDCCH in congestion rate(%)],
[SDCCH DROP Rate %],
[SDCCH assignments success rate%],
[TCH in call drop rate(include handover)(%)],
[TCH dropped call total number],
[Call drop rate of traffic],
[Handover success rate(%)],
[Handover in success rate(%)],
[Handover out success rate(%)],
[TCH attempt total number(include handover)],
[TCH overflow total number(include handover)],
[TCH seizure total number(include handover)],
[Call Set-up Success Rate(CSSR)(%)],
[Rx_Quality DL (0-5)],
[Rx_Quality UL (0-5)],
[TA=0 %],
[0<TA<=2 %],
[2<TA<=4 %],
[4<TA<=8 %],
[8<TA<=20 %],
[20<TA<=63 %],
[TA>63 %] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$] order by [Begin time]", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }




                        int Ericsson_CC2 = 0; //1
                        int Huawei_CC2 = 0;  //2
                        int Nokia_CC2 =0;  //3
                        int Ericsson_CC3 = 1; //1
                        int Huawei_CC3 = 0;  //2
                        int Nokia_CC3 = 0;  //3
                        int Ericsson_RD3 = 1;  //1
                        int Huawei_RD3 =0;  //2
                        int Nokia_RD3 = 0;  //3 
                        int Ericsson_RD4 = 0;  //1
                        int Huawei_RD4 =0;  //3
                        int Nokia_RD4 = 0;  //2


                        
                        // 2G
                        if (Ericsson_CC2 == 1 && Table_Name == "CC2 Eric Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [BSC], [CELL], [REGION], [PROVINCE], [Date], [CSSR_MCI], [OHSR], [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] INTO [CC2_TBL_Ericsson]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [CC2_TBL_Ericsson]";
                            string IMPORT_STR_5 = string.Format(@" select [BSC], [CELL], [REGION], [PROVINCE], [Date], [CSSR_MCI], [OHSR], [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        if (Huawei_CC2 == 1 && Table_Name == "CC2 Huawei Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [BSC], [CELL], [REGION], [PROVINCE], [Date], [CSSR3], [OHSR2] , [CDR3] INTO [CC2_TBL_Huawei]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [CC2_TBL_Huawei]";
                            string IMPORT_STR_5 = string.Format(@" select [BSC], [CELL], [REGION], [PROVINCE], [Date], [CSSR3], [OHSR2] , [CDR3] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Nokia_CC2 == 1 && Table_Name == "CC2 Nokia SEG Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [BSC], [SEG], [REGION], [PROVINCE], [Date], [CSSR_MCI], [OHSR], [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)] INTO [CC2_TBL_Nokia]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [CC2_TBL_Nokia]";
                            string IMPORT_STR_5 = string.Format(@" select [BSC], [SEG], [REGION], [PROVINCE], [Date], [CSSR_MCI], [OHSR], [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }




                        //if (Ericsson_CC3 == 1 && Table_Name == "Ericsson CC3 Cell Daily")
                        //{
                        //    ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                        //    connection = new SqlConnection(ConnectionString);
                        //    connection.Open();

                        //    string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Date], [Cs_RAB_Establish_Success_Rate], [CS_Drop_Call_Rate] INTO [CC3_TBL_Ericsson]";
                        //    string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                        //    Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                        //    // Other Orders to Fill Table
                        //    string IMPORT_STR_4 = "INSERT INTO [CC3_TBL_Ericsson]";
                        //    string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Date], [Cs_RAB_Establish_Success_Rate], [CS_Drop_Call_Rate] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                        //    Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        //}

                        if (Ericsson_CC3 == 1 && Table_Name == "Ericsson CC3 Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Date], [CS_Traffic] INTO [CC3_TBL_Voice]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [CC3_TBL_Voice]";
                            string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Date], [CS_Traffic] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Huawei_CC3==1 && Table_Name == "Huawei CC3 Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Date], [CS_RAB_Setup_Success_Ratio], [AMR_Call_Drop_Ratio_New(Hu_CELL)] INTO [CC3_TBL_Huawei]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [CC3_TBL_Huawei]";
                            string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Date], [CS_RAB_Setup_Success_Ratio], [AMR_Call_Drop_Ratio_New(Hu_CELL)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                       }


                        if (Nokia_CC3 == 1 && Table_Name == "Nokia CC3 Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Date], [CS_RAB_Establish_Success_Rate], [CS_Drop_Call_Rate] INTO [CC3_TBL_Nokia]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [CC3_TBL_Nokia]";
                            string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Date], [CS_RAB_Establish_Success_Rate], [CS_Drop_Call_Rate] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        //if (Ericsson_RD3 == 1 && Table_Name == "Ericsson RD3 Cell Daily")
                        //{
                        //    ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                        //    connection = new SqlConnection(ConnectionString);
                        //    connection.Open();

                        //    string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Province], [Date], [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)], [HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)] INTO [RD3_TBL_Ericsson]";
                        //    string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                        //    Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                        //    // Other Orders to Fill Table
                        //    string IMPORT_STR_4 = "INSERT INTO [RD3_TBL_Ericsson]";
                        //    string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Province], [Date], [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)], [HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                        //    Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        //}



                        if (Ericsson_RD3 == 1 && Table_Name == "Ericsson RD3 Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Province], [Date], [PS_Volume(GB)(UCell_Eric)] INTO [RD3_TBL_Data]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [RD3_TBL_Data]";
                            string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Province], [Date], [PS_Volume(GB)(UCell_Eric)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        if (Huawei_RD3 == 1 && Table_Name == "Huawei RD3 Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Province], [Date], [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)], [HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)] INTO [RD3_TBL_Huawei]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [RD3_TBL_Huawei]";
                            string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Province], [Date], [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)], [HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Nokia_RD3 == 1 && Table_Name == "Nokia RD3 Cell Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [ElementID], [ElementID1], [Province], [Date], [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)], [Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)] INTO [RD3_TBL_Nokia]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [RD3_TBL_Nokia]";
                            string IMPORT_STR_5 = string.Format(@" select [ElementID], [ElementID1], [Province], [Date], [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)], [Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }




                        // 4G
                        if (Ericsson_RD4 == 1 && Table_Name == "Ericsson Cell Info Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select  [Datetime], [eNodeB], [Index], [Province], [Region], [E_RAB_Drop_Rate(eNodeB_Eric)], [Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)], [Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)] INTO [RD4_TBL_Ericsson]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [RD4_TBL_Ericsson]";
                            string IMPORT_STR_5 = string.Format(@" select  [Datetime], [eNodeB],  [Index], [Province], [Region], [E_RAB_Drop_Rate(eNodeB_Eric)], [Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)], [Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Datetime", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Huawei_RD4 == 1 && Table_Name == "Huawei Cell Info Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Datetime], [eNodeB], [Index], [Province], [Region], [Call_Drop_Rate], [Average_Downlink_User_Throughput(Mbit/s)], [Average_UPlink_User_Throughput(Mbit/s)] INTO [RD4_TBL_Huawei]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [RD4_TBL_Huawei]";
                            string IMPORT_STR_5 = string.Format(@" select [Datetime], [eNodeB], [Index], [Province], [Region], [Call_Drop_Rate], [Average_Downlink_User_Throughput(Mbit/s)], [Average_UPlink_User_Throughput(Mbit/s)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Datetime", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }

                        if (Nokia_RD4 == 1 && Table_Name == "Nokia Cell Info Daily")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=RAN; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], [ElementID1], [Index], [Province], [Region], [E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)], [User_Throughput_DL_mbps(Nokia_LTE_CELL)], [User_Throughput_UL_mbps(Nokia_LTE_CELL)] INTO [RD4_TBL_Nokia]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [RD4_TBL_Nokia]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], [ElementID1], [Index], [Province], [Region], [E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)], [User_Throughput_DL_mbps(Nokia_LTE_CELL)], [User_Throughput_UL_mbps(Nokia_LTE_CELL)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }






                        // Medical Tables
                        if (Table_Name == "3G KPI")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Medical; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], [Site], [Sector], [Cell], [DL_Power_Utilization], [CS_RAB_Congestion], [CS_RRC_SR], [HS_USER_Throughput], [PS_RAB_Congestion], [Soft_HO_SR], [Total_Data_Volume], [Voice_Drop_Call_Rate], [Voice_Traffic] INTO [Medical_3G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Medical_3G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], [Site], [Sector], [Cell], [DL_Power_Utilization], [CS_RAB_Congestion], [CS_RRC_SR], [HS_USER_Throughput], [PS_RAB_Congestion], [Soft_HO_SR], [Total_Data_Volume], [Voice_Drop_Call_Rate], [Voice_Traffic] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        // Medical Tables
                        if (Table_Name == "2G KPI")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Medical; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], [BSC], [Site], [Sector], [SDCCH_Cong], [SDCCH_Traffic], [CDR], [CSSR], [IHSR], [OHSR], [TCH_ASFR], [TCH_Availability], [TCH_Cong], [TCH_Traffic(Erlang)] INTO [Medical_2G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Medical_2G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], [BSC], [Site], [Sector], [SDCCH_Cong], [SDCCH_Traffic], [CDR], [CSSR], [IHSR], [OHSR], [TCH_ASFR], [TCH_Availability], [TCH_Cong], [TCH_Traffic(Erlang)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        // Medical Tables
                        if (Table_Name == "4G KPI")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Medical; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], [Site], [Sector], [NE], [RRC_Connected_Users], [UE_DL_THR(Mbps)], [DL_PRB_Utilization_Rate], [LTE_Service_Setup_SR], [Total_Volume(UL+DL)(GB)] INTO [Medical_4G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Medical_4G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], [Site], [Sector], [NE], [RRC_Connected_Users], [UE_DL_THR(Mbps)], [DL_PRB_Utilization_Rate], [LTE_Service_Setup_SR], [Total_Volume(UL+DL)(GB)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        //Tehran Dashboards
                        if (Table_Name == "Ericsson 2G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Eric_BSC)], 	[Cell_Availability(Eric_BSC)], 	[CSSR_MCI(Congestion_included)(Eric_BSC)], 	[IHSR(Eric_BSC)], 	[OHSR(Eric_BSC)], 	[Payload_Total(TB)(Eric_BSC)], 	[RxQual_DL(Eric_BSC)], 	[RxQual_UL(Eric_BSC)], 	[SDCCH_Access_Succ_Rate_New(Eric_BSC)], 	[SDCCH_Congestion_Rate(Eric_BSC)], 	[SDCCH_Drop_Rate(Eric_BSC)], 	[TCH_Assign_Fail_Rate(Congestion_Excluded)(Eric_BSC)], 	[TCH_Congestion_Rate(Eric_BSC)], 	[TCH_Traffic(Erlang)(Eric_BSC)] INTO [Ericsson_2G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Ericsson_2G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Eric_BSC)], 	[Cell_Availability(Eric_BSC)], 	[CSSR_MCI(Congestion_included)(Eric_BSC)], 	[IHSR(Eric_BSC)], 	[OHSR(Eric_BSC)], 	[Payload_Total(TB)(Eric_BSC)], 	[RxQual_DL(Eric_BSC)], 	[RxQual_UL(Eric_BSC)], 	[SDCCH_Access_Succ_Rate_New(Eric_BSC)], 	[SDCCH_Congestion_Rate(Eric_BSC)], 	[SDCCH_Drop_Rate(Eric_BSC)], 	[TCH_Assign_Fail_Rate(Congestion_Excluded)(Eric_BSC)], 	[TCH_Congestion_Rate(Eric_BSC)], 	[TCH_Traffic(Erlang)(Eric_BSC)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        if (Table_Name == "Huawei 2G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(HU_BSC)], 	[TCH_Availability(HU_BSC)], 	[CSSR3(HU_BSC)], 	[IHSR2(HU_BSC)], 	[OHSR2(HU_BSC)], 	[Payload_Total_TB(HU_BSC)], 	[RX_QUALITTY_DL_NEW(HUAWEI_BSC)], 	[RX_QUALITTY_UL_NEW(HUAWEI_BSC)], 	[SDCCH_Access_Success_Rate2(HU_BSC)], 	[SDCCH_Congestion_Rate(HU_BSC)], 	[SDCCH_Drop_Rate(HU_BSC)], 	[TCH_Assignment_FR(HU_BSC)], 	[TCH_Cong(HU_BSC)], 	[TCH_Traffic(HU_BSC)]  INTO[Huawei_2G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Huawei_2G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(HU_BSC)], 	[TCH_Availability(HU_BSC)], 	[CSSR3(HU_BSC)], 	[IHSR2(HU_BSC)], 	[OHSR2(HU_BSC)], 	[Payload_Total_TB(HU_BSC)], 	[RX_QUALITTY_DL_NEW(HUAWEI_BSC)], 	[RX_QUALITTY_UL_NEW(HUAWEI_BSC)], 	[SDCCH_Access_Success_Rate2(HU_BSC)], 	[SDCCH_Congestion_Rate(HU_BSC)], 	[SDCCH_Drop_Rate(HU_BSC)], 	[TCH_Assignment_FR(HU_BSC)], 	[TCH_Cong(HU_BSC)], 	[TCH_Traffic(HU_BSC)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Table_Name == "Nokia 2G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Nokia_BSC)], 	[TCH_Availability(Nokia_BSC)], 	[CSSR_MCI(Nokia_BSC)], 	[IHSR(Nokia_BSC)], 	[OHSR(Nokia_BSC)], 	[Payload_Data(UL+DL)_TB(Nokia_BSC)], 	[RxQuality_DL(Nokia_BSC)], 	[RxQuality_UL(Nokia_BSC)], 	[SDCCH_Access_Success_Rate(Nokia_BSC)], 	[SDCCH_Congestion_Rate(Nokia_BSC)], 	[SDCCH_Drop_Rate(Nokia_BSC)], 	[TCH_Assignment_FR(Nokia_BSC)], 	[TCH_Cong_Rate(Nokia_BSC)], 	[TCH_Traffic(Nokia_BSC)]  INTO[Nokia_2G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Nokia_2G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Nokia_BSC)], 	[TCH_Availability(Nokia_BSC)], 	[CSSR_MCI(Nokia_BSC)], 	[IHSR(Nokia_BSC)], 	[OHSR(Nokia_BSC)], 	[Payload_Data(UL+DL)_TB(Nokia_BSC)], 	[RxQuality_DL(Nokia_BSC)], 	[RxQuality_UL(Nokia_BSC)], 	[SDCCH_Access_Success_Rate(Nokia_BSC)], 	[SDCCH_Congestion_Rate(Nokia_BSC)], 	[SDCCH_Drop_Rate(Nokia_BSC)], 	[TCH_Assignment_FR(Nokia_BSC)], 	[TCH_Cong_Rate(Nokia_BSC)], 	[TCH_Traffic(Nokia_BSC)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }




                        if (Table_Name == "Ericsson 2G Hourly")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Eric_BSC)], 	[TCH_Availability(Eric_BSC)], 	[CSSR_MCI(Congestion_included)(Eric_BSC)], 	[IHSR(Eric_BSC)], 	[OHSR(Eric_BSC)], 	[Payload_Total(TB)(Eric_BSC)], 	[RxQual_DL(Eric_BSC)], 	[RxQual_UL(Eric_BSC)], 	[SDCCH_Access_Succ_Rate_New(Eric_BSC)], 	[SDCCH_Congestion_Rate(Eric_BSC)], 	[SDCCH_Drop_Rate(Eric_BSC)], 	[TCH_Assign_Fail_Rate(Congestion_Excluded)(Eric_BSC)], 	[TCH_Congestion_Rate(Eric_BSC)], 	[TCH_Traffic(Erlang)(Eric_BSC)] INTO [Ericsson_2G_Hourly]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Ericsson_2G_Hourly]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Eric_BSC)], 	[TCH_Availability(Eric_BSC)], 	[CSSR_MCI(Congestion_included)(Eric_BSC)], 	[IHSR(Eric_BSC)], 	[OHSR(Eric_BSC)], 	[Payload_Total(TB)(Eric_BSC)], 	[RxQual_DL(Eric_BSC)], 	[RxQual_UL(Eric_BSC)], 	[SDCCH_Access_Succ_Rate_New(Eric_BSC)], 	[SDCCH_Congestion_Rate(Eric_BSC)], 	[SDCCH_Drop_Rate(Eric_BSC)], 	[TCH_Assign_Fail_Rate(Congestion_Excluded)(Eric_BSC)], 	[TCH_Congestion_Rate(Eric_BSC)], 	[TCH_Traffic(Erlang)(Eric_BSC)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        if (Table_Name == "Huawei 2G Hourly")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(HU_BSC)], 	[TCH_Availability(HU_BSC)], 	[CSSR3(HU_BSC)], 	[IHSR2(HU_BSC)], 	[OHSR2(HU_BSC)], 	[Payload_Total_TB(HU_BSC)], 	[RX_QUALITTY_DL_NEW(HUAWEI_BSC)], 	[RX_QUALITTY_UL_NEW(HUAWEI_BSC)], 	[SDCCH_Access_Success_Rate2(HU_BSC)], 	[SDCCH_Congestion_Rate(HU_BSC)], 	[SDCCH_Drop_Rate(HU_BSC)], 	[TCH_Assignment_FR(HU_BSC)], 	[TCH_Cong(HU_BSC)], 	[TCH_Traffic(HU_BSC)]  INTO[Huawei_2G_Hourly]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Huawei_2G_Hourly]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(HU_BSC)], 	[TCH_Availability(HU_BSC)], 	[CSSR3(HU_BSC)], 	[IHSR2(HU_BSC)], 	[OHSR2(HU_BSC)], 	[Payload_Total_TB(HU_BSC)], 	[RX_QUALITTY_DL_NEW(HUAWEI_BSC)], 	[RX_QUALITTY_UL_NEW(HUAWEI_BSC)], 	[SDCCH_Access_Success_Rate2(HU_BSC)], 	[SDCCH_Congestion_Rate(HU_BSC)], 	[SDCCH_Drop_Rate(HU_BSC)], 	[TCH_Assignment_FR(HU_BSC)], 	[TCH_Cong(HU_BSC)], 	[TCH_Traffic(HU_BSC)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Table_Name == "Nokia 2G Hourly")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Nokia_BSC)], 	[TCH_Availability(Nokia_BSC)], 	[CSSR_MCI(Nokia_BSC)], 	[IHSR(Nokia_BSC)], 	[OHSR(Nokia_BSC)], 	[Payload_Data(UL+DL)_TB(Nokia_BSC)], 	[RxQuality_DL(Nokia_BSC)], 	[RxQuality_UL(Nokia_BSC)], 	[SDCCH_Access_Success_Rate(Nokia_BSC)], 	[SDCCH_Congestion_Rate(Nokia_BSC)], 	[SDCCH_Drop_Rate(Nokia_BSC)], 	[TCH_Assignment_FR(Nokia_BSC)], 	[TCH_Cong_Rate(Nokia_BSC)], 	[TCH_Traffic(Nokia_BSC)]  INTO[Nokia_2G_Hourly]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Nokia_2G_Hourly]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[2G_Voice_Call_Drop_Rate(Nokia_BSC)], 	[TCH_Availability(Nokia_BSC)], 	[CSSR_MCI(Nokia_BSC)], 	[IHSR(Nokia_BSC)], 	[OHSR(Nokia_BSC)], 	[Payload_Data(UL+DL)_TB(Nokia_BSC)], 	[RxQuality_DL(Nokia_BSC)], 	[RxQuality_UL(Nokia_BSC)], 	[SDCCH_Access_Success_Rate(Nokia_BSC)], 	[SDCCH_Congestion_Rate(Nokia_BSC)], 	[SDCCH_Drop_Rate(Nokia_BSC)], 	[TCH_Assignment_FR(Nokia_BSC)], 	[TCH_Cong_Rate(Nokia_BSC)], 	[TCH_Traffic(Nokia_BSC)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        if (Table_Name == "Ericsson 3G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[Cell_Availability_Rate_Include_Blocking(RNC_Eric)], 	[CS_IRAT_HO_Suc_Rate(RNC_Eric)], 	[CS_Setup_Success_Rate(RNC_Eric)], 	[HS_USER_Throughput_NET_PQ(Mbps)(RNC_Eric)], 	[HSDPA_Scheduling_Cell_Throughput(Kbps)(RNC_Eric)], 	[IFHO_Success_Rate(RNC_Eric)], 	[PS_Drop_Call_Rate(RNC_Eric)], 	[PS_Setup_Success_Rate(RNC_Eric)], 	[PS_Volume(TB)(RNC_Eric)], 	[Soft_HO_Suc_Rate(RNC_Eric)], 	[uplink_average_RSSI(dbm)(RNC_Eric)], 	[Voice_Drop_Call_Rate(RNC_Eric)], 	[Voice_Traffic(RNC_Eric)] INTO[Ericsson_3G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Ericsson_3G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[Cell_Availability_Rate_Include_Blocking(RNC_Eric)], 	[CS_IRAT_HO_Suc_Rate(RNC_Eric)], 	[CS_Setup_Success_Rate(RNC_Eric)], 	[HS_USER_Throughput_NET_PQ(Mbps)(RNC_Eric)], 	[HSDPA_Scheduling_Cell_Throughput(Kbps)(RNC_Eric)], 	[IFHO_Success_Rate(RNC_Eric)], 	[PS_Drop_Call_Rate(RNC_Eric)], 	[PS_Setup_Success_Rate(RNC_Eric)], 	[PS_Volume(TB)(RNC_Eric)], 	[Soft_HO_Suc_Rate(RNC_Eric)], 	[uplink_average_RSSI(dbm)(RNC_Eric)], 	[Voice_Drop_Call_Rate(RNC_Eric)], 	[Voice_Traffic(RNC_Eric)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        if (Table_Name == "Huawei 3G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[Radio_Network_Availability_Ratio(Hu_RNC)], 	[CS_IRAT_HO_SR(Hu_RNC)], 	[CS_CSSR(Hu_RNC)], 	[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(RNC_HUAWEI)], 	[HSDPA_SCHEDULING_CELL_THROUGHPUT(Mbit/s)(RNC_HUAWEI)], 	[Inter_frequency_Hard_Handover_Success_Rate(Hu_RNC)], 	[PS_Call_Drop_Ratio(Hu_RNC)_NEW], 	[PS_CSSR_Hu], 	[PS_Total_payload(TB)(Hu_RNC)], 	[Soft_Handover_Succ_Rate(Hu_RNC)], 	[average_RTWP_dbm(RNC_Hu)], 	[AMR_Call_Drop_Ratio_New(Hu_RNC)], 	[3G_VOICE_TRAFFIC(Huawei_RNC)]  INTO[Huawei_3G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Huawei_3G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[Radio_Network_Availability_Ratio(Hu_RNC)], 	[CS_IRAT_HO_SR(Hu_RNC)], 	[CS_CSSR(Hu_RNC)], 	[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(RNC_HUAWEI)], 	[HSDPA_SCHEDULING_CELL_THROUGHPUT(Mbit/s)(RNC_HUAWEI)], 	[Inter_frequency_Hard_Handover_Success_Rate(Hu_RNC)], 	[PS_Call_Drop_Ratio(Hu_RNC)_NEW], 	[PS_CSSR_Hu], 	[PS_Total_payload(TB)(Hu_RNC)], 	[Soft_Handover_Succ_Rate(Hu_RNC)], 	[average_RTWP_dbm(RNC_Hu)], 	[AMR_Call_Drop_Ratio_New(Hu_RNC)], 	[3G_VOICE_TRAFFIC(Huawei_RNC)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Table_Name == "Nokia 3G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[Cell_Availability_including_blocked_by_user_state_RNC], 	[CS_IRAT_hardhandover_2GTO3G(RNC_NOKIA)], 	[CS_RRCSETUP_SR(Nokia_RNC)], 	[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(RNC_NOKIA)], 	[Active_HS-DSCH_cell_throughput_mbs(RNC_nokia)], 	[Intra_RNC_Inter_frequency_HO_Success_Rate_RT(RNC)], 	[PS_RAB_drop_rate(RNC)], 	[Packet_Session_stp_SR_RNC], 	[PS_Payload_Total(HS+R99)(Nokia_RNC)_TB], 	[Soft_HO_Success_rate_RT(RNC)], 	[RTWP(dbm)_New(RNC_NOKIA)], 	[CS_Drop_Call_Rate], 	[3G_VOICE_TRAFFIC(RNC_nokia)]  INTO[Nokia_3G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Nokia_3G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[Cell_Availability_including_blocked_by_user_state_RNC], 	[CS_IRAT_hardhandover_2GTO3G(RNC_NOKIA)], 	[CS_RRCSETUP_SR(Nokia_RNC)], 	[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(RNC_NOKIA)], 	[Active_HS-DSCH_cell_throughput_mbs(RNC_nokia)], 	[Intra_RNC_Inter_frequency_HO_Success_Rate_RT(RNC)], 	[PS_RAB_drop_rate(RNC)], 	[Packet_Session_stp_SR_RNC], 	[PS_Payload_Total(HS+R99)(Nokia_RNC)_TB], 	[Soft_HO_Success_rate_RT(RNC)], 	[RTWP(dbm)_New(RNC_NOKIA)], 	[CS_Drop_Call_Rate], 	[3G_VOICE_TRAFFIC(RNC_nokia)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }





                        if (Table_Name == "Ericsson 4G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[Average_PDCP_Cell_Dl_Throughput(Mbps)(Prov_EricLTE)], 	[Average_UE_DL_Latency(ms)(Prov_EricLTE)], 	[Average_UE_DL_Throughput(Mbps)(Prov_EricLTE)], 	[Cell_Availability_Rate_Include_Blocking(Prov_EricLTE)], 	[CQI(Prov_EricLTE)], 	[CSFB_Success_Rate(Prov_EricLTE)], 	[E_RAB_Drop_Rate(Prov_EricLTE)], 	[E-RAB_Setup_SR_incl_added_New(Prov_EricLTE)], 	[InterF_Handover_Execution_Rate(Prov_EricLTE)], 	[IntraF_Handover_Execution_Rate(Prov_EricLTE)], 	[LTE_Service_Setup_SR(Prov_EricLTE)],  [RSSI_PUCCH(Prov_EricLTE)], [RSSI_PUSCH(Prov_EricLTE)],   [RRC_Connection_Setup_Success_Rate(Prov_EricLTE)], 	[S1Signal_Estab_Success_Rate(Prov_EricLTE)], 	[Total_Traffic(UL+DL)(TB)(Prov_EricLTE)], 	[VoLTE_Traffic_Erlang_QCI1(Prov_EricLTE)] INTO[Ericsson_4G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Ericsson_4G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[Average_PDCP_Cell_Dl_Throughput(Mbps)(Prov_EricLTE)], 	[Average_UE_DL_Latency(ms)(Prov_EricLTE)], 	[Average_UE_DL_Throughput(Mbps)(Prov_EricLTE)], 	[Cell_Availability_Rate_Include_Blocking(Prov_EricLTE)], 	[CQI(Prov_EricLTE)], 	[CSFB_Success_Rate(Prov_EricLTE)], 	[E_RAB_Drop_Rate(Prov_EricLTE)], 	[E-RAB_Setup_SR_incl_added_New(Prov_EricLTE)], 	[InterF_Handover_Execution_Rate(Prov_EricLTE)], 	[IntraF_Handover_Execution_Rate(Prov_EricLTE)], 	[LTE_Service_Setup_SR(Prov_EricLTE)], [RSSI_PUCCH(Prov_EricLTE)], [RSSI_PUSCH(Prov_EricLTE)],   [RRC_Connection_Setup_Success_Rate(Prov_EricLTE)], 	[S1Signal_Estab_Success_Rate(Prov_EricLTE)], 	[Total_Traffic(UL+DL)(TB)(Prov_EricLTE)], 	[VoLTE_Traffic_Erlang_QCI1(Prov_EricLTE)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }



                        if (Table_Name == "Huawei 4G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[Downlink_Cell_Throghput(Mbit/s)(Prov_HuLTE)], 	[Average_DL_Latency_ms(Prov_HuLTE)], 	[Average_Downlink_User_Throghput(Mbit/s)(Prov_HuLTE)], 	[Cell_Availability_Rate_include_Blocking(Prov_HuLTE)], 	[Average_CQI(Prov_HuLTE)], 	[CSFB_Success_Rate(Prov_HuLTE)], 	[E-RAB_Drop_New(Hu_LTE_Prov)], 	[E-RAB_Setup_Success_Rate_Prov(Prov_HuLTE)], 	[InterF_HOOut_SR_Prov], 	[IntraF_HOOut_SR_Prov], 	[CSSR(ALL)(Hu_Prov)], [RSSI_PUCCH(Huawei_PROV)], [RSSI_PUSCH(Huawei_PROV)],	[RRC_Connection_Setup_Success_Rate_service_Prov], 	[S1signal_Connection_Setup_SR(Huawei_LTE_Prov)], 	[Total_Traffic(TB)_Prov], 	[Volte_Traffic_Erlang(Prov_HuLTE)]  INTO[Huawei_4G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Huawei_4G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[Downlink_Cell_Throghput(Mbit/s)(Prov_HuLTE)], 	[Average_DL_Latency_ms(Prov_HuLTE)], 	[Average_Downlink_User_Throghput(Mbit/s)(Prov_HuLTE)], 	[Cell_Availability_Rate_include_Blocking(Prov_HuLTE)], 	[Average_CQI(Prov_HuLTE)], 	[CSFB_Success_Rate(Prov_HuLTE)], 	[E-RAB_Drop_New(Hu_LTE_Prov)], 	[E-RAB_Setup_Success_Rate_Prov(Prov_HuLTE)], 	[InterF_HOOut_SR_Prov], 	[IntraF_HOOut_SR_Prov], 	[CSSR(ALL)(Hu_Prov)], [RSSI_PUCCH(Huawei_PROV)],  [RSSI_PUSCH(Huawei_PROV)],	[RRC_Connection_Setup_Success_Rate_service_Prov], 	[S1signal_Connection_Setup_SR(Huawei_LTE_Prov)], 	[Total_Traffic(TB)_Prov], 	[Volte_Traffic_Erlang(Prov_HuLTE)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        if (Table_Name == "Nokia 4G")
                        {
                            ConnectionString = @"Server=NAKPRG-NB1243; Database=Dashboards; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();

                            string IMPORT_STR_1 = "select [Date], 	[NE], 	[PDCP_Layer_Active_Cell_Throughput_DL_mbps(Nokia_LTE_Prov)], 	[Average_Latency_DL_ms(Nokia_LTE_Prov)], 	[User_Throughput_DL_mbps(Nokia_LTE_Prov)], 	[cell_availability_include_manual_blocking(Nokia_LTE_Prov)], 	[Average_CQI(Nokia_LTE_Prov)], 	[Init_Contx_stp_SR_for_CSFB(Totonchi)], 	[E-RAB_Drop_New(Nokia_LTE_Prov)], 	[E-RAB_Setup_SR_incl_added(Nokia_LTE_Prov)], 	[Inter-Freq_HO_SR(Nokia_LTE_Prov)], 	[Intra-Freq_HO_Success_Ratio(Nokia_LTE_Prov)], 	[Initial_E-RAB_Accessibility(Nokia_LTE_Prov)], [Average_RSSI_for_PUCCH(Nokia_LTE_Prov)], [Average_RSSI_for_PUSCH(Nokia_LTE_Prov)],	[RRC_Connection_Setup_Success_Ratio(Nokia_LTE_Prov)], 	[S1Signal_E-RAB_Setup_SR(Nokia_LTE_Prov)], 	[Total_Payload_TB(Nokia_LTE_Prov)], 	[Total_Voice_Traffic_QCI1_Erlang(Nokia_LTE_Prov)]  INTO[Nokia_4G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Nokia_4G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[NE], 	[PDCP_Layer_Active_Cell_Throughput_DL_mbps(Nokia_LTE_Prov)], 	[Average_Latency_DL_ms(Nokia_LTE_Prov)], 	[User_Throughput_DL_mbps(Nokia_LTE_Prov)], 	[cell_availability_include_manual_blocking(Nokia_LTE_Prov)], 	[Average_CQI(Nokia_LTE_Prov)], 	[Init_Contx_stp_SR_for_CSFB(Totonchi)], 	[E-RAB_Drop_New(Nokia_LTE_Prov)], 	[E-RAB_Setup_SR_incl_added(Nokia_LTE_Prov)], 	[Inter-Freq_HO_SR(Nokia_LTE_Prov)], 	[Intra-Freq_HO_Success_Ratio(Nokia_LTE_Prov)], 	[Initial_E-RAB_Accessibility(Nokia_LTE_Prov)], [Average_RSSI_for_PUCCH(Nokia_LTE_Prov)], [Average_RSSI_for_PUSCH(Nokia_LTE_Prov)],	[RRC_Connection_Setup_Success_Ratio(Nokia_LTE_Prov)], 	[S1Signal_E-RAB_Setup_SR(Nokia_LTE_Prov)], 	[Total_Payload_TB(Nokia_LTE_Prov)], 	[Total_Voice_Traffic_QCI1_Erlang(Nokia_LTE_Prov)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                            Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        }


                        // 2G Contractual WPC
                        //if (Table_Name == "NAK-Tehran" || Table_Name == "NAK-Huawei" || Table_Name == "NAK-Nokia" || Table_Name == "NAK-North" || Table_Name == "NAK-Alborz")
                        //{
                        //    //ConnectionString = @"Server=NAKPRG-NB1243; Database=Contract; Trusted_Connection=True;";
                        //    //connection = new SqlConnection(ConnectionString);
                        //    //connection.Open();

                        //    ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
                        //    connection = new SqlConnection(ConnectionString);
                        //    connection.Open();


                        //    string IMPORT_STR_1 = "select [Date], 	[Cell], 	[BSC], 	[Vendor], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency Index], 	[Status], 	[QIX], [QIxP], 	[QIxBL], 	[AVG_TCH_Traffic], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of CSSR], 	[Worst(%) of OHSR], 	[Worst(%) of CDR], 	[Worst(%) of TCH_Assignment_FR], 	[Worst(%) of DL Quality <=4], 	[Worst(%) of UL Quality <=4], 	[Worst(%) of SDCCH_Congestion_Rate], 	[Worst(%) of SDCCH_Access_Success_Rate], 	[Worst(%) of SDCCH_Drop_Rate], 	[Worst(%) of IHSR], 	[Worst(%) of AMRHR_Usage]  INTO [Contractual_WPC_2G]";
                        //    string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                        //    Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                        //    // Other Orders to Fill Table
                        //    string IMPORT_STR_4 = "INSERT INTO [Contractual_WPC_2G]";
                        //    string IMPORT_STR_5 = string.Format(@" select [Date], 	[Cell], 	[BSC], 	[Vendor], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency Index], 	[Status], [QIX],	[QIxP], 	[QIxBL], 	[AVG_TCH_Traffic], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of CSSR], 	[Worst(%) of OHSR], 	[Worst(%) of CDR], 	[Worst(%) of TCH_Assignment_FR], 	[Worst(%) of DL Quality <=4], 	[Worst(%) of UL Quality <=4], 	[Worst(%) of SDCCH_Congestion_Rate], 	[Worst(%) of SDCCH_Access_Success_Rate], 	[Worst(%) of SDCCH_Drop_Rate], 	[Worst(%) of IHSR], 	[Worst(%) of AMRHR_Usage] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                        //    Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        //}

                        // 3G CS Contractual WPC
                        //if (Table_Name == "NAK-Tehran" || Table_Name == "NAK-Huawei" || Table_Name == "NAK-Nokia" || Table_Name == "NAK-North" || Table_Name == "NAK-Alborz")
                        //{
                        //    ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
                        //    connection = new SqlConnection(ConnectionString);
                        //    connection.Open();


                        //    string IMPORT_STR_1 = "select [Date], 	[Cell], 	[RNC], 	[Vendor], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency index], 	[Status], 	[QIxP], [QIx],	[QIxBL], 	[AVG_CS_Traffic)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (CS)], 	[Worst(%) of W2G IRAT/IF HO success rate], 	[Worst(%) of Drop Call Rate], 	[Worst(%) of Soft HO Success Rate], 	[Worst(%) of CS RRC Connection Establishment SR (%)]  INTO [Contractual_WPC_3G_CS]";
                        //    string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                        //    Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                        //    // Other Orders to Fill Table
                        //    string IMPORT_STR_4 = "INSERT INTO [Contractual_WPC_3G_CS]";
                        //    string IMPORT_STR_5 = string.Format(@" select [Date], 	[Cell], 	[RNC], 	[Vendor], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency index], 	[Status], [QIx],	[QIxP], 	[QIxBL], 	[AVG_CS_Traffic)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (CS)], 	[Worst(%) of W2G IRAT/IF HO success rate], 	[Worst(%) of Drop Call Rate], 	[Worst(%) of Soft HO Success Rate], 	[Worst(%) of CS RRC Connection Establishment SR (%)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                        //    Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        //}

                        //3G PS Contractual WPC
                        //if (Table_Name == "NAK-Tehran" || Table_Name == "NAK-Huawei" || Table_Name == "NAK-Nokia" || Table_Name == "NAK-North" || Table_Name == "NAK-Alborz")
                        //{
                        //    ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
                        //    connection = new SqlConnection(ConnectionString);
                        //    connection.Open();


                        //    string IMPORT_STR_1 = "select  [Date], 	[Cell], 	[RNC], 	[Vendor], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency Index], 	[Status], [QIx],	[QIxP], 	[QIxBL], 	[Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	[Worst(%) of RAB Establishment  Success Rate (EUL)], 	[Worst(%) of EUL MAC User Throughput (kbps)], 	[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	[Worst(%) of RAB Drop Rate (HSDPA)], 	[Worst(%) of RAB Drop Rate (EUL)], 	[Worst(%) of MultiRAB Setup Success Ratio (%)], 	[Worst(%) of PS_RRC_Setup_Success_Rate], 	[Worst(%) of Ps_RAB_Establish_Success_Rate], 	[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	[Worst(%) of Drop_Call_Rate], 	[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	[Worst(%) of HS_share_PAYLOAD_%], 	[Worst(%) of HSDPA Cell Throughput (Mbps)] INTO [Contractual_WPC_3G_PS]";
                        //    string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                        //    Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                        //    // Other Orders to Fill Table
                        //    string IMPORT_STR_4 = "INSERT INTO [Contractual_WPC_3G_PS]";
                        //    string IMPORT_STR_5 = string.Format(@" select [Date], 	[Cell], 	[RNC], 	[Vendor], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency Index], 	[Status], [QIx],	[QIxP], 	[QIxBL], 	[Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	[Worst(%) of RAB Establishment  Success Rate (EUL)], 	[Worst(%) of EUL MAC User Throughput (kbps)], 	[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	[Worst(%) of RAB Drop Rate (HSDPA)], 	[Worst(%) of RAB Drop Rate (EUL)], 	[Worst(%) of MultiRAB Setup Success Ratio (%)], 	[Worst(%) of PS_RRC_Setup_Success_Rate], 	[Worst(%) of Ps_RAB_Establish_Success_Rate], 	[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	[Worst(%) of Drop_Call_Rate], 	[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	[Worst(%) of HS_share_PAYLOAD_%], 	[Worst(%) of HSDPA Cell Throughput (Mbps)] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                        //    string IMPORT_STR_6 = Table_Name + string.Format(@"$'] order by Date", filePath);
                        //    Import_S_Second = IMPORT_STR_4 + IMPORT_STR_5 + IMPORT_STR_6;
                        //}

                        // 4G Contractual WPC
                        if (Table_Name == "NAK-Tehran" || Table_Name == "NAK-Huawei" || Table_Name == "NAK-Nokia" || Table_Name == "NAK-North" || Table_Name == "NAK-Alborz")
                        {
                            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
                            connection = new SqlConnection(ConnectionString);
                            connection.Open();


                            string IMPORT_STR_1 = "select [Date], 	[eNodeB], [RNC], 	[Vendor], 	[province], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency Index], 	[Status], 	[QIx], 	[QIxP], 	[QIxBL], 	[Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RRC Connection Establishment Success Rate], 	[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	[Worst(%) of ERAB Stablishment Success Rate (Added)], 	[Worst(%) of DL User Troughput  (Mbps)], 	[Worst(%) of UL User Throughput (Mbps)], 	[Worst(%) of Handover Success Rate], 	[Worst(%) of ERAB Drop rate], 	[Worst(%) of UE Context Drop Rate], 	[Worst(%) of S1 Signalling Success Rate], 	[Worst(%) of Inter Frequency Handover Execution SR (%)], 	[Worst(%) of Intra Frequency Handover Execution SR (%)], 	[Worst(%) of Average Ul Packet Loss Rate (%)], 	[Worst(%) of payload per Carrier] INTO [Contractual_WPC_4G]";
                            string IMPORT_STR_2 = string.Format(@" from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
                            string IMPORT_STR_3 = Table_Name + string.Format(@"$']", filePath);
                            Import_S_First = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;

                            // Other Orders to Fill Table
                            string IMPORT_STR_4 = "INSERT INTO [Contractual_WPC_4G]";
                            string IMPORT_STR_5 = string.Format(@" select [Date], 	[eNodeB], [RNC],	[Vendor], 	[province], 	[Contractor], 	[LEVEL], 	[Worst], 	[Effeciency Index], 	[Status], 	[QIx], 	[QIxP], 	[QIxBL], 	[Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RRC Connection Establishment Success Rate], 	[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	[Worst(%) of ERAB Stablishment Success Rate (Added)], 	[Worst(%) of DL User Troughput  (Mbps)], 	[Worst(%) of UL User Throughput (Mbps)], 	[Worst(%) of Handover Success Rate], 	[Worst(%) of ERAB Drop rate], 	[Worst(%) of UE Context Drop Rate], 	[Worst(%) of S1 Signalling Success Rate], 	[Worst(%) of Inter Frequency Handover Execution SR (%)], 	[Worst(%) of Intra Frequency Handover Execution SR (%)], 	[Worst(%) of Average Ul Packet Loss Rate (%)], 	[Worst(%) of payload per Carrier] from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...['", filePath);
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
                            if ((Name_Table.Rows[k]).ItemArray[0].ToString() == Table_Name)
                            {
                                table_finder++;
                                SqlCommand Import_command = new SqlCommand(Import_S_Second, connection);
                                Import_command.ExecuteNonQuery();
                                break;
                            }
                            if ((Name_Table.Rows[k]).ItemArray[0].ToString() == "Medical_2G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Medical_4G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Medical_3G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "CC3_TBL_Ericsson" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "CC3_TBL_Huawei" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "CC3_TBL_Nokia" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "RD3_TBL_Ericsson" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "RD3_TBL_Huawei" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "RD3_TBL_Nokia" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "CC3_TBL_Voice" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "RD3_TBL_Data" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Ericsson_2G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Huawei_2G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Nokia_2G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Ericsson_3G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Huawei_3G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Nokia_3G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Ericsson_4G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Huawei_4G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Nokia_4G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Contractual_WPC_2G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Contractual_WPC_3G_CS" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Contractual_WPC_3G_PS" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Contractual_WPC_4G" || (Name_Table.Rows[k]).ItemArray[0].ToString() == "Ericsson_2G_Hourly")
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