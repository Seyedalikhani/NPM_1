using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
//using DevExtreme.NETCore.Demos.Models;
//using DevExtreme.NETCore.Demos.Models.DataGrid;
//using DevExtreme.NETCore.Demos.Models.SampleData;
//using Microsoft.AspNetCore.Mvc;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;



namespace NPM_1.Controllers
{

    public class MAPController : Controller
    {
        // GET: MAP
        //public ActionResult Index()
        //{
        //    return View();
        //}


        public ActionResult MAP()

        {
            // Loads the initial view and populates dropdowns with provinces and corresponding cities.
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Provinces
                SqlCommand cmd1 = new SqlCommand("SELECT DISTINCT Province_EN as 'Province' FROM ARAS_DB order by Province_EN", conn);
                SqlDataAdapter da1 = new SqlDataAdapter(cmd1);
                DataTable dt1 = new DataTable();
                da1.Fill(dt1);
                ViewBag.ProvinceList = dt1.AsEnumerable().Select(r => r[0].ToString()).ToList();

                string firstProvince = ViewBag.ProvinceList[0];
                SqlCommand cmd2 = new SqlCommand("SELECT DISTINCT Location as 'Site' FROM ARAS_DB WHERE Province_EN=@p", conn);
                cmd2.Parameters.AddWithValue("@p", firstProvince);
                SqlDataAdapter da2 = new SqlDataAdapter(cmd2);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                ViewBag.SiteList = dt2.AsEnumerable().Select(r => r[0].ToString()).ToList();



            }



            return View();
        }



        [HttpPost]
        public JsonResult GetLocations(string selected_province)
        {
            List<string> sites = new List<string>();
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT DISTINCT Location FROM ARAS_DB WHERE Province_EN = @p", conn);
                cmd.Parameters.AddWithValue("@p", selected_province);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    sites.Add(reader.GetString(0));
                }
            }

            return Json(sites);
        }



    }
}