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
using System.Drawing;

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
            return View(new List<ExcelUpload>());
        }

        [HttpPost]
        public ActionResult ExcelUpload(HttpPostedFileBase file)
        {
            List<ExcelUpload> data = new List<ExcelUpload>();

            bool useRealFileName = true;
            bool excelHaveHeader = false;
            bool excelRange = true;
            string _path = "~/Content/";
            string _fileName = useRealFileName ? Path.GetFileName(file.FileName) : Guid.NewGuid().ToString();
            string _savePath = Path.Combine(Server.MapPath(_path), _fileName);

            using (var reader = new BinaryReader(file.InputStream))
            {
                file.SaveAs(_savePath);
            }

            var excel = new ExcelQueryFactory(_savePath);

            // Query Worksheet Names
            var worksheetNames = excel.GetWorksheetNames();

            // Query Column Name
            //by default, the queried worksheet name is Sheet1. To query different worksheet pass the worksheet name or worksheet index(worksheet is ordered alphabeticly) in like below example. 
            var columnNames = excel.GetColumnNames("Student Data");

            #region Property to Column Mapping 
            //(if the column name is exactly same with the property name you can skip this region)
            excel.AddMapping("StudentId", "Student ID");
            // or
            excel.AddMapping<ExcelUpload>(x => x.StudentName , "Student Name");
            #endregion

            if (excelHaveHeader)
            {
                if(!excelRange)
                { 
                    var excelContent = from a in excel.Worksheet<ExcelUpload>("Student Data")
                                        select a;
                }
                else
                {
                    //Data from only a specific range of cells within a worksheet can be queried as well.
                    var excelContent = from a in excel.WorksheetRange<ExcelUpload>("A1","B12","Student Data")
                                    select a;

                    //var excelContents = from a in excel.WorksheetRange<ExcelUpload>("D1", "E12")
                    //                    select a;
                }
            }
            else
            {
                // query a worksheet without header row
                if (!excelRange)
                {
                    var excelContent = from a in excel.WorksheetNoHeader(0) //Queries the first worksheet in alphabetical order
                                       select a;
                }
                else
                {
                    var excelContent = from a in excel.WorksheetRangeNoHeader("A2", "B12","Student Data")
                                       select a;
                }
            }



            #region Using the LinqToExcel.Row class
            //Query results can be returned as LinqToExcel.Row objects which allows you to access a cell's value by using the column name in the string index. Just use the Worksheet() method without a generic argument.

            var excelContentAccessByColumnName = from c in excel.Worksheet()
                                   where c["Student ID"] == "0" || c["Student Name"] == "John"
                                   select c;

            var excelContentCast = from c in excel.Worksheet()
                                 where c["Student ID"].Cast<int>() == 0
                                 select c;

            #endregion

            data = (from a in excel.Worksheet<ExcelUpload>("Student Data")
                   select a).ToList();

            // for further information you can read the full documentation in : https://github.com/paulyoder/LinqToExcel

            return View(data);
        }

        public ActionResult ExcelDownload()
        {
            List<ExcelUpload> dummylist = new List<ExcelUpload>();

            for(int i = 0; i < 11; i++)
            {
                dummylist.Add(new ExcelUpload { StudentId = i, StudentName = "John" + i.ToString() });
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

                    // To Add Named Range
                    worksheet.Names.Add("testnamerange", worksheet.Cells["A1:B1"]);

                    int row = 1;
                    foreach(var item in dummylist)
                    {
                        row++;
                        worksheet.Cells["A" + row].Value = item.StudentId;
                        worksheet.Cells["B" + row].Value = item.StudentName;
                    }

                    row++;
                    worksheet.Cells["A"+row+":B"+row].Merge = true;
                    worksheet.Cells["A" + row + ":B" + row].Value = "END";
                    worksheet.Cells["A" + row + ":B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["A1:B" + row].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                    worksheet.Cells["A1:B" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A1:B" + row].Style.Border.Right.Style = ExcelBorderStyle.DashDotDot;
                    worksheet.Cells["A1:B" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                    Image logo = Image.FromFile(Path.Combine(Server.MapPath("~/Content/"), "1.png"));

                    //Below for resize the image
                    

                    var picture = worksheet.Drawings.AddPicture("imagename", logo);
                    picture.SetPosition(13, 0, 0, 0);
                    picture.SetSize(50);
                    //picture.SetSize(300,500);

                    byteArray = package.GetAsByteArray();
                }
            }
            
            return File(new MemoryStream(byteArray, 0, byteArray.Length), "application/octet-stream", "ExcelDownload.xlsx");
        }
    }
}