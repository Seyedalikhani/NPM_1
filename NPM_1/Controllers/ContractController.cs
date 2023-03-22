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



namespace NPM_1.Controllers
{
    public class ContractController : Controller
    {
        // GET: AZWLL
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult Contract_KPI()
        {

      



      

            return View();
        }




        [HttpPost]
        public JsonResult dropdown_kpiPost_Cell(string text)
        {
            //bool flag = true;
            //string responseMessage = string.Empty;
            //string interval_province_name = text;
            ////  return Json(new { success = true, result = text });

            //string ConnectionString = @"Server=NAKPRG-NB1243\AHMAD; Database=AZWLL; Trusted_Connection=True;";
            //SqlConnection connection = new SqlConnection(ConnectionString);
            //connection.Open();

            //string Cell_STR_1 = "";
            //if (interval_province_name == "Daily/West Azarbaijan")
            //{
            //    Cell_STR_1 = "select distinct [UserLabel] from [KPI_AG_Daily]";
            //}
            //if (interval_province_name == "BH/West Azarbaijan")
            //{
            //    Cell_STR_1 = "select distinct [UserLabel] from [KPI_AG_BH]";
            //}
            //if (interval_province_name == "Daily/East Azarbaijan")
            //{
            //    Cell_STR_1 = "select distinct [Bts] from [KPI_AS_Daily]";
            //}
            //if (interval_province_name == "BH/East Azarbaijan")
            //{
            //    Cell_STR_1 = "select distinct [Bts] from [KPI_AS_BH]";
            //}

            //SqlCommand Cell_Command1 = new SqlCommand(Cell_STR_1, connection);
            //Cell_Command1.ExecuteNonQuery();

            //DataTable Cell_Table = new DataTable();
            //SqlDataAdapter dataAdapter = new SqlDataAdapter(Cell_Command1);
            //dataAdapter.Fill(Cell_Table);

            //List<Cell> Cell_LIST = new List<Cell>();
            //string Cell_Name = "";
            //for (int k = 1; k <= Cell_Table.Rows.Count; k++)
            //{
            //    Cell_Name = (Cell_Table.Rows[k - 1]).ItemArray[0].ToString();
            //    Cell_LIST.Add(new Cell
            //    {
            //        CellName = Cell_Name
            //    });
            //}

            return Json("kl;jlkk", JsonRequestBehavior.AllowGet);
        }






    }
}