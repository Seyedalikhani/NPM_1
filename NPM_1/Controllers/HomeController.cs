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
    public class HomeController : Controller
    {
        public ActionResult Index()
        {

            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            //SqlConnection connection = new SqlConnection(ConnectionString);
            //connection.Open();

            //string Cell_Select = "select * from Country_Traffic order by Date";
            //SqlCommand Cell_Select1 = new SqlCommand(Cell_Select, connection);
            //Cell_Select1.ExecuteNonQuery();

            //DataTable Cell_Select_Table = new DataTable();
            //SqlDataAdapter dataAdapter = new SqlDataAdapter(Cell_Select1);
            //dataAdapter.Fill(Cell_Select_Table);

            //List<DataPoint> dataPoints = new List<DataPoint>();

            //for (int k = 0; k <= Cell_Select_Table.Rows.Count - 1; k++)
            //{
            //    string date = Cell_Select_Table.Rows[k].ItemArray[0].ToString();
            //    //DateTime dt = Convert.ToDateTime(date);
            //    //double dt1 = dt.Ticks;
            //    string kpi = Cell_Select_Table.Rows[k].ItemArray[1].ToString();
            //    double kpi1 = Convert.ToDouble(kpi);
            //    // dataPoints.Add(new DataPoint(dt1, kpi1));

            //    TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
            //    dataPoints.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1));
            //}

            //ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoints);


            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            //SqlConnection connection = new SqlConnection(ConnectionString);
            //connection.Open();

            //string Cell_Select = "select * from Country_Traffic order by Date";
            //SqlCommand Cell_Select1 = new SqlCommand(Cell_Select, connection);
            //Cell_Select1.ExecuteNonQuery();

            //DataTable Cell_Select_Table = new DataTable();
            //SqlDataAdapter dataAdapter = new SqlDataAdapter(Cell_Select1);
            //dataAdapter.Fill(Cell_Select_Table);

            //List<DataPoint> dataPoints = new List<DataPoint>();
            //List<DataPoint> dataPoints2 = new List<DataPoint>();

            //for (int k = 0; k <= Cell_Select_Table.Rows.Count - 1; k++)
            //{
            //    string date = Cell_Select_Table.Rows[k].ItemArray[0].ToString();
            //    //DateTime dt = Convert.ToDateTime(date);
            //    //double dt1 = dt.Ticks;
            //    string Traffic_kpi = Cell_Select_Table.Rows[k].ItemArray[1].ToString();
            //    double kpi1 = Convert.ToDouble(Traffic_kpi);

            //    string Data_kpi = Cell_Select_Table.Rows[k].ItemArray[2].ToString();
            //    double kpi2 = Convert.ToDouble(Data_kpi);

            //    // dataPoints.Add(new DataPoint(dt1, kpi1));

            //    TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
            //    dataPoints.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1));
            //    dataPoints2.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi2));
            //}



            //ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoints);

            return View();



        }




        public ContentResult JSON()
        {

            string Server_Name = @"AHMAD\" + "SQLEXPRESS";
            string DataBase_Name = "NAK";

            string ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";

            //string ConnectionString = @"Server=NAKPRG-NB1243; Database=Data; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();

            string Country_Traffic_str = "select * from Country_Traffic order by Date";
            SqlCommand Country_Traffic_sql = new SqlCommand(Country_Traffic_str, connection);
            Country_Traffic_sql.ExecuteNonQuery();

            DataTable Country_Traffic_Table = new DataTable();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(Country_Traffic_sql);
            dataAdapter.Fill(Country_Traffic_Table);

            List<DataPoint> dataPoints1 = new List<DataPoint>();

            for (int k = 0; k <= Country_Traffic_Table.Rows.Count - 1; k++)
            {
                string date = Country_Traffic_Table.Rows[k].ItemArray[0].ToString();

                string Traffic_kpi = Country_Traffic_Table.Rows[k].ItemArray[1].ToString();
                double kpi1 = Convert.ToDouble(Traffic_kpi);

                string Data_kpi = Country_Traffic_Table.Rows[k].ItemArray[2].ToString();
                double kpi2 = Convert.ToDouble(Data_kpi);

                TimeSpan ts1 = DateTime.Parse(date) - DateTime.Parse("1970-01-01 00:00");
                dataPoints1.Add(new DataPoint(Math.Truncate(ts1.TotalMilliseconds), kpi1,kpi2));

            }

            //ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoints);
            //ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoints2);

   

            JsonSerializerSettings _jsonSetting = new JsonSerializerSettings() { NullValueHandling = NullValueHandling.Ignore };
            return Content(JsonConvert.SerializeObject(dataPoints1, _jsonSetting), "application/json");

    
        }






        public class DataPoint
        {
            public DataPoint(double x, double y, double z)
            {
                this.x = x;
                this.y = y;
                this.z = z;
            }

 
            public Nullable<double> x = null;
            public Nullable<double> y = null;
            public Nullable<double> z = null;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        public ActionResult Reports()
        {
            return View();
        }

        public ActionResult Contract()
        {
            return View();
        }


        public ActionResult TXPO()
        {
            return View();
        }

        public ActionResult WPC()
        {
            return View();
        }
    }
}