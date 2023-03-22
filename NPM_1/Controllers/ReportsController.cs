using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace NPM_1.Controllers
{
    public class ReportsController : Controller
    {
        // GET: Reports
        public ActionResult NUR()
        {
            return View();
        }

        
        public ActionResult Delete_Old_Data()
        {

            string ConnectionString = @"Server=NAKPRG-NB1243\AHMAD; Database=NUR; Trusted_Connection=True;";
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();
            SqlCommand cc1 = new SqlCommand("truncate table [dbo].[Ericsson_BSC_1]", connection);
            cc1.ExecuteNonQuery();
            SqlCommand cc2 = new SqlCommand("truncate table [dbo].[Ericsson_BSC_2]", connection);
            cc2.ExecuteNonQuery();
            SqlCommand cc3 = new SqlCommand("truncate table [dbo].[Ericsson_BSC_3]", connection);
            cc3.ExecuteNonQuery();
            SqlCommand cc4 = new SqlCommand("truncate table [dbo].[Ericsson_BSC_4]", connection);
            cc4.ExecuteNonQuery();
            SqlCommand cc5 = new SqlCommand("truncate table [dbo].[Ericsson_BSC_5]", connection);
            cc5.ExecuteNonQuery();
            SqlCommand cc6 = new SqlCommand("truncate table [dbo].[Tehran_CR_1]", connection);
            cc6.ExecuteNonQuery();
            SqlCommand cc7 = new SqlCommand("truncate table [dbo].[Tehran_CR_2]", connection);
            cc7.ExecuteNonQuery();
            SqlCommand cc8 = new SqlCommand("truncate table [dbo].[Tehran_CR_3]", connection);
            cc8.ExecuteNonQuery();
            SqlCommand cc9 = new SqlCommand("truncate table [dbo].[Tehran_CR_4]", connection);
            cc9.ExecuteNonQuery();
            SqlCommand cc10 = new SqlCommand("truncate table [dbo].[Tehran_CR_5]", connection);
            cc10.ExecuteNonQuery();

            SqlCommand cc11 = new SqlCommand("truncate table [dbo].[Ericsson_BSC]", connection);
            cc11.ExecuteNonQuery();
            SqlCommand cc12 = new SqlCommand("truncate table [dbo].[Tehran_CR]", connection);
            cc12.ExecuteNonQuery();
            SqlCommand cc13 = new SqlCommand("truncate table [dbo].[Cells_DowonTime]", connection);
            cc13.ExecuteNonQuery();
            SqlCommand cc14 = new SqlCommand("truncate table [dbo].[Availability]", connection);
            cc14.ExecuteNonQuery();
            SqlCommand cc15 = new SqlCommand("truncate table [dbo].[Availbility_DateSite]", connection);
            cc15.ExecuteNonQuery();
            SqlCommand cc16 = new SqlCommand("truncate table [dbo].[Cells_DowonTime]", connection);
            cc16.ExecuteNonQuery();
            SqlCommand cc17 = new SqlCommand("truncate table [dbo].[Tehran_CR_Avail_Flag]", connection);
            cc17.ExecuteNonQuery();
            SqlCommand cc18 = new SqlCommand("truncate table [dbo].[Tehran_CR_Avail_Issue]", connection);
            cc18.ExecuteNonQuery();

            return View();
        }


   

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
                if (file != null && file.ContentLength > 0 && (Path.GetExtension(file.FileName).ToLower() == ".xlsx" || Path.GetExtension(file.FileName).ToLower() == ".xls"))
                {
                    try
                    {
                        string fileName = Path.GetFileName(file.FileName);
                        string filePath = Path.Combine(Server.MapPath("~/UploadFiles"), fileName);
                        file.SaveAs(filePath);

                        flag = true;
                        responseMessage = "Upload Successful.";
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












        public ActionResult SMS()
        {
            return View();
        }
    }
}