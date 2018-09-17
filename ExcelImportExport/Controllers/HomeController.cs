using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelImportExport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
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

        public ActionResult ExcelUpload()
        {
            //List<ViewHouseDataUploadValidation> Validate = new List<ViewHouseDataUploadValidation>();
            //return View(new ViewHouseDataUpload { Validate = Validate, error = null });
            return View();
        }
    }
}