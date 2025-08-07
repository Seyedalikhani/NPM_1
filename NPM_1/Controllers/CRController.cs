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
    public class CRController : Controller
    {
        // GET: CR
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult CR_KPI()
        {


           return View();
        }

    }
}