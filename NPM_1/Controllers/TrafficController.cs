using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.InkML;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Helpers;
using System.Web.Mvc;

namespace NPM_1.Controllers
{
    public class TrafficController : Controller
    {
        public ActionResult Traffic_KPI()    // (GET)
        {
            // Loads the initial view and populates dropdowns with provinces and corresponding cities.
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Provinces
                SqlCommand cmd1 = new SqlCommand("SELECT DISTINCT Province FROM SA_Province_Cont_MAP where Province!='Iran'", conn);
                SqlDataAdapter da1 = new SqlDataAdapter(cmd1);
                DataTable dt1 = new DataTable();
                da1.Fill(dt1);
                ViewBag.ProvinceList = dt1.AsEnumerable().Select(r => r[0].ToString()).ToList();

                // Default cities for first province
                string firstProvince = ViewBag.ProvinceList[0];
                SqlCommand cmd2 = new SqlCommand("SELECT DISTINCT City FROM ARAS_Coverage WHERE Province=@p AND City IS NOT NULL", conn);
                cmd2.Parameters.AddWithValue("@p", firstProvince);
                SqlDataAdapter da2 = new SqlDataAdapter(cmd2);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                ViewBag.CityList = dt2.AsEnumerable().Select(r => r[0].ToString()).ToList();



                // Default sites for first province
                SqlCommand cmd3 = new SqlCommand("SELECT Site FROM [SA_Site_Traffic_Sharing] WHERE Province=@p and Datetime = (SELECT MAX(Datetime) FROM [SA_Site_Traffic_Sharing]) order by Site", conn);
                cmd3.Parameters.AddWithValue("@p", firstProvince);
                SqlDataAdapter da3 = new SqlDataAdapter(cmd3);
                DataTable dt3 = new DataTable();
                da3.Fill(dt3);
                ViewBag.SiteList = dt3.AsEnumerable().Select(r => r[0].ToString()).ToList();

            }

            return View();
        }

        // AJAX handler to dynamically load cities when a province is selected.
        // Input: Province name.
        // Output: List of city names as JSON.
        [HttpPost]
        public JsonResult GetCities(string selected_province)          // (POST)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            List<string> cities = new List<string>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT DISTINCT City FROM ARAS_Coverage WHERE Province = @province AND City IS NOT NULL";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@province", selected_province);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    cities.Add(reader.GetString(0));
                }
            }

            return Json(cities);
        }



        // AJAX handler to dynamically load sites when a province is selected.
        // Input: Province name.
        // Output: List of site names as JSON.
        [HttpPost]
        public JsonResult GetSites(string selected_province)          // (POST)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            List<string> sites = new List<string>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT Site FROM [SA_Site_Traffic_Sharing] WHERE Province=@province and Datetime = (SELECT MAX(Datetime) FROM [SA_Site_Traffic_Sharing]) order by Site";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@province", selected_province);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    sites.Add(reader.GetString(0));
                }
            }

            return Json(sites);
        }





        // AJAX handler to retrieve traffic data (voice & data) based on selected province.
        // Query:
        //From SA_Site_Traffic_Sharing.
        //Filters numeric values.
        //Aggregates by Datetime.
        //Output: List of DataPoint objects (with x = timestamp, y = Erlang, z = GB), returned as JSON.

        [HttpPost]
        public JsonResult GetProvinceTrafficData(string province)    // (POST)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            List<DataPoint> dataPoints = new List<DataPoint>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = @"
                SELECT [Datetime], sum(cast([Traffic_Erlang] as float)) as 'Traffic (Erlang)', sum(cast([Payload_GB] as float)) as 'Payload (GB)' 
                FROM [SA_Site_Traffic_Sharing]
                WHERE Province = @province 
                and ISNUMERIC([Traffic_Erlang]) = 1 and ISNUMERIC([Payload_GB]) = 1
                group by Datetime
                ORDER BY Datetime";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@province", province);
                SqlDataReader reader = cmd.ExecuteReader();


                while (reader.Read())
                {
                    DateTime dt = Convert.ToDateTime(reader["Datetime"]);
                    double voice = Convert.ToDouble(reader["Traffic (Erlang)"]);
                    double data = Convert.ToDouble(reader["Payload (GB)"]);
                    double x = (dt - new DateTime(1970, 1, 1)).TotalMilliseconds;
                    dataPoints.Add(new DataPoint(x, voice, data));
                }
            }

            return Json(dataPoints);
        }


        // AJAX handler to retrieve traffic data (voice & data) based on selected province & city.
        // Query:
        //From SA_Site_Traffic_Sharing.
        //Filters numeric values.
        //Aggregates by Datetime.
        //Output: List of DataPoint objects (with x = timestamp, y = Erlang, z = GB), returned as JSON.

        [HttpPost]
        public JsonResult GetCityTrafficData(string province, string city)    // (POST)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            List<DataPoint> dataPoints = new List<DataPoint>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = @"
                SELECT [Datetime], sum(cast([Traffic_Erlang] as float)) as 'Traffic (Erlang)', sum(cast([Payload_GB] as float)) as 'Payload (GB)' 
                FROM [SA_Site_Traffic_Sharing]
                WHERE Province = @province and City=@city
                and ISNUMERIC([Traffic_Erlang]) = 1 and ISNUMERIC([Payload_GB]) = 1
                group by Datetime
                ORDER BY Datetime";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@province", province);
                cmd.Parameters.AddWithValue("@city", city);
                SqlDataReader reader = cmd.ExecuteReader();


                while (reader.Read())
                {
                    DateTime dt = Convert.ToDateTime(reader["Datetime"]);
                    double voice = Convert.ToDouble(reader["Traffic (Erlang)"]);
                    double data = Convert.ToDouble(reader["Payload (GB)"]);
                    double x = (dt - new DateTime(1970, 1, 1)).TotalMilliseconds;
                    dataPoints.Add(new DataPoint(x, voice, data));
                }
            }

            return Json(dataPoints);
        }



        // AJAX handler to retrieve traffic sharing (voice & data) based on selected province
        [HttpPost]
        public JsonResult GetProvinceTrafficShare(string province)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            double payload3G = 0, payload4G = 0, payload5G = 0, totalPayload = 0;
            double traffic2G = 0, traffic3G = 0, traffic4G = 0, totalTraffic = 0;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = @"
              SELECT 
                SUM(ISNULL(TRY_CAST([_3G_Payload_GB] AS FLOAT), 0)) AS Payload3G,
                SUM(ISNULL(TRY_CAST([_4G_Payload_GB] AS FLOAT), 0)) AS Payload4G,
                SUM(ISNULL(TRY_CAST([_5G_Payload_GB] AS FLOAT), 0)) AS Payload5G,
                SUM(ISNULL(TRY_CAST([Payload_GB] AS FLOAT), 0)) AS TotalPayload,
                SUM(ISNULL(TRY_CAST([_2G_Traffic_Erlang] AS FLOAT), 0)) AS Traffic2G,
                SUM(ISNULL(TRY_CAST([_3G_Traffic_Erlang] AS FLOAT), 0)) AS Traffic3G,
                SUM(ISNULL(TRY_CAST([_4G_Traffic_Erlang] AS FLOAT), 0)) AS Traffic4G,
                SUM(ISNULL(TRY_CAST([Traffic_Erlang] AS FLOAT), 0)) AS TotalTraffic
            FROM [SA_Site_Traffic_Sharing]
            WHERE Province = @province AND Datetime = (SELECT MAX(Datetime) FROM [SA_Site_Traffic_Sharing])";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@province", province);


                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    payload3G = reader["Payload3G"] != DBNull.Value ? Convert.ToDouble(reader["Payload3G"]) : 0;
                    payload4G = reader["Payload4G"] != DBNull.Value ? Convert.ToDouble(reader["Payload4G"]) : 0;
                    payload5G = reader["Payload5G"] != DBNull.Value ? Convert.ToDouble(reader["Payload5G"]) : 0;
                    totalPayload = reader["TotalPayload"] != DBNull.Value ? Convert.ToDouble(reader["TotalPayload"]) : 0;

                    traffic2G = reader["Traffic2G"] != DBNull.Value ? Convert.ToDouble(reader["Traffic2G"]) : 0;
                    traffic3G = reader["Traffic3G"] != DBNull.Value ? Convert.ToDouble(reader["Traffic3G"]) : 0;
                    traffic4G = reader["Traffic4G"] != DBNull.Value ? Convert.ToDouble(reader["Traffic4G"]) : 0;
                    totalTraffic = reader["TotalTraffic"] != DBNull.Value ? Convert.ToDouble(reader["TotalTraffic"]) : 0;
                }
            }

            var result = new
            {
                Payload3GPercent = totalPayload > 0 ? (payload3G / totalPayload) * 100 : 0,
                Payload4GPercent = totalPayload > 0 ? (payload4G / totalPayload) * 100 : 0,
                Payload5GPercent = totalPayload > 0 ? (payload5G / totalPayload) * 100 : 0,
                Traffic2GPercent = totalTraffic > 0 ? (traffic2G / totalTraffic) * 100 : 0,
                Traffic3GPercent = totalTraffic > 0 ? (traffic3G / totalTraffic) * 100 : 0,
                Traffic4GPercent = totalTraffic > 0 ? (traffic4G / totalTraffic) * 100 : 0
            };

            return Json(result);
        }




        // AJAX handler to retrieve traffic sharing (voice & data) based on selected province & city
        [HttpPost]
        public JsonResult GetCityTrafficShare(string province, string city)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            double payload3G = 0, payload4G = 0, payload5G = 0, totalPayload = 0;
            double traffic2G = 0, traffic3G = 0, traffic4G = 0, totalTraffic = 0;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = @"
              SELECT 
                SUM(ISNULL(TRY_CAST([_3G_Payload_GB] AS FLOAT), 0)) AS Payload3G,
                SUM(ISNULL(TRY_CAST([_4G_Payload_GB] AS FLOAT), 0)) AS Payload4G,
                SUM(ISNULL(TRY_CAST([_5G_Payload_GB] AS FLOAT), 0)) AS Payload5G,
                SUM(ISNULL(TRY_CAST([Payload_GB] AS FLOAT), 0)) AS TotalPayload,
                SUM(ISNULL(TRY_CAST([_2G_Traffic_Erlang] AS FLOAT), 0)) AS Traffic2G,
                SUM(ISNULL(TRY_CAST([_3G_Traffic_Erlang] AS FLOAT), 0)) AS Traffic3G,
                SUM(ISNULL(TRY_CAST([_4G_Traffic_Erlang] AS FLOAT), 0)) AS Traffic4G,
                SUM(ISNULL(TRY_CAST([Traffic_Erlang] AS FLOAT), 0)) AS TotalTraffic
            FROM [SA_Site_Traffic_Sharing]
            WHERE Province = @province AND City = @city AND Datetime = (SELECT MAX(Datetime) FROM [SA_Site_Traffic_Sharing])";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@province", province);
                cmd.Parameters.AddWithValue("@city", city);


                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    payload3G = reader["Payload3G"] != DBNull.Value ? Convert.ToDouble(reader["Payload3G"]) : 0;
                    payload4G = reader["Payload4G"] != DBNull.Value ? Convert.ToDouble(reader["Payload4G"]) : 0;
                    payload5G = reader["Payload5G"] != DBNull.Value ? Convert.ToDouble(reader["Payload5G"]) : 0;
                    totalPayload = reader["TotalPayload"] != DBNull.Value ? Convert.ToDouble(reader["TotalPayload"]) : 0;

                    traffic2G = reader["Traffic2G"] != DBNull.Value ? Convert.ToDouble(reader["Traffic2G"]) : 0;
                    traffic3G = reader["Traffic3G"] != DBNull.Value ? Convert.ToDouble(reader["Traffic3G"]) : 0;
                    traffic4G = reader["Traffic4G"] != DBNull.Value ? Convert.ToDouble(reader["Traffic4G"]) : 0;
                    totalTraffic = reader["TotalTraffic"] != DBNull.Value ? Convert.ToDouble(reader["TotalTraffic"]) : 0;
                }
            }

            var result = new
            {
                Payload3GPercent = totalPayload > 0 ? (payload3G / totalPayload) * 100 : 0,
                Payload4GPercent = totalPayload > 0 ? (payload4G / totalPayload) * 100 : 0,
                Payload5GPercent = totalPayload > 0 ? (payload5G / totalPayload) * 100 : 0,
                Traffic2GPercent = totalTraffic > 0 ? (traffic2G / totalTraffic) * 100 : 0,
                Traffic3GPercent = totalTraffic > 0 ? (traffic3G / totalTraffic) * 100 : 0,
                Traffic4GPercent = totalTraffic > 0 ? (traffic4G / totalTraffic) * 100 : 0
            };

            return Json(result);
        }




        // Used to encapsulate traffic data for plotting.
        public class DataPoint
        {
            public double x { get; set; }
            public double y { get; set; }
            public double z { get; set; }

            public DataPoint(double x, double y, double z)
            {
                this.x = x;
                this.y = y;
                this.z = z;
            }
        }
    }

}
