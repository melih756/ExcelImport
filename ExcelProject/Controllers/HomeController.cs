using ExcelProject.Models;
using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelProject.Controllers
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

        public ActionResult CreateExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var person = new List<Person>
            {
                new Person { Id = 1, Name = "Melih" },
                new Person { Id = 2, Name = "Ahmet" }
            };
            using (var stream = new MemoryStream())
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, "person.xlsx");

                using (var xlPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = xlPackage.Workbook.Worksheets.Add("Person2");
                    worksheet.Cells["A1"].Value = "personeller";
                    using (var r = worksheet.Cells["A1:C1"])
                    {
                        r.Merge = true;
                        r.Style.Font.Color.SetColor(Color.White);
                        r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                        r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                    }
                    worksheet.Cells["A4"].Value = "ID";
                    worksheet.Cells["B4"].Value = "AD SOYAD";
                    worksheet.Cells["A4:C4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells["A4:C4"].Style.Fill.BackgroundColor.SetColor(Color.Azure);
                    worksheet.Cells["A4:C4"].Style.Font.Bold = true;

                    int row = 5;
                    foreach (var excelModel in person)
                    {
                        worksheet.Cells[row, 1].Value = excelModel.Id;
                        worksheet.Cells[row, 2].Value = excelModel.Name;
                        row++;
                    }

                    xlPackage.Save();
                    return View();
                }

            }

        }
        public ActionResult ReadExcelFile()
        {
            string path = @"C:\Desktop\IIS Express\wwwroot";
            string filePath = Path.Combine(path, "person.xlsx");
            FileInfo fi = new FileInfo(path);

            ExcelPackage excelpackage = new ExcelPackage(fi);
            ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.FirstOrDefault();

            int rows = worksheet.Dimension.Rows;
            int column = worksheet.Dimension.Columns;

            var person = new List<Person>();
            

            return View();
        }

    }
}
