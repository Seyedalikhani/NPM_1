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
    public class WPCController : Controller
    {
        public ActionResult WPC_KPI()    // (GET)
        {
            // Loads the initial view and populates dropdowns with provinces and corresponding cities.
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Provinces
                SqlCommand cmd1 = new SqlCommand("SELECT DISTINCT Province FROM SA_Province_Cont_MAP", conn);
                SqlDataAdapter da1 = new SqlDataAdapter(cmd1);
                DataTable dt1 = new DataTable();
                da1.Fill(dt1);
                ViewBag.ProvinceList = dt1.AsEnumerable().Select(r => r[0].ToString()).ToList();

             

            }

            return View();
        }






        //[HttpPost]
        //public JsonResult FetchFilteredData(List<string> filters)
        //{
        //    // Example: Parse your filters and query DB
        //    // You'll likely need a better model structure than just List<string>

        //    var results = new List<YourDataModel>();

        //    foreach (var f in filters)
        //    {
        //        // Parse or match "Province: X, Technology: Y, Date: ..., ..." and query your DB
        //        // This is pseudo-code:
        //        var parsed = ParseFilterString(f); // write this method
        //        var query = db.YourTable.Where(x =>
        //            x.Province == parsed.Province &&
        //            x.Technology == parsed.Technology &&
        //            x.Date == parsed.Date &&
        //            x.Interval == parsed.Interval &&
        //            x.KPI == parsed.KPI
        //        );

        //        results.AddRange(query);
        //    }

        //    return Json(results);
        //}






    }
}
