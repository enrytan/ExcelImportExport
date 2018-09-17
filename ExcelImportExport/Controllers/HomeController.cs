using ExcelImportExport.Models;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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

        public ActionResult ExcelDownload()
        {
            List<ExcelUpload> dummylist = new List<ExcelUpload>();

            for(int i = 0; i < 11; i++)
            {
                dummylist.Add(new ExcelUpload { StudentId = i, StudenteName = "John" + i.ToString() });
            }

            byte[] byteArray;

            using (MemoryStream mem = new MemoryStream())
            {
                using (ExcelPackage package = new ExcelPackage(mem))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Student Data");

                    worksheet.Cells["A1"].Value = "Student ID";
                    worksheet.Cells["B1"].Value = "Student Name";
                    worksheet.Column(1).Width = 30;
                    worksheet.Column(2).Width = 30;
                    worksheet.Cells["A1:B1"].Style.Font.Bold = true;

                    int row = 1;
                    foreach(var item in dummylist)
                    {
                        row++;
                        worksheet.Cells["A" + row].Value = item.StudentId;
                        worksheet.Cells["B" + row].Value = item.StudenteName;
                    }

                    row++;
                    worksheet.Cells["A"+row+":B"+row].Merge = true;
                    worksheet.Cells["A" + row + ":B" + row].Value = "END";
                    worksheet.Cells["A" + row + ":B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["A1:B" + row].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                    worksheet.Cells["A1:B" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A1:B" + row].Style.Border.Right.Style = ExcelBorderStyle.DashDotDot;
                    worksheet.Cells["A1:B" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                    byteArray = package.GetAsByteArray();
                }
            }
            
            return File(new MemoryStream(byteArray, 0, byteArray.Length), "application/octet-stream", "ExcelDownload.xlsx");
        }
    }
}