using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using SampleExcel.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace SampleExcel.Controllers
{
    public class UserController : Controller
    {
        // GET: UserController
        public ActionResult Index()
        {
            var users = GetUsers();
            return View(users);
        }

        private List<User> GetUsers()
        {
            var users = new List<User>()
            {
                new User(){ Name = "alex", Email = "alex@test.fr", Phone = "8676875976"},
                new User(){ Name = "seb", Email = "seb@test.fr", Phone = "8676959867"},
                new User(){ Name = "kate", Email = "kate@test.fr", Phone = "345677808"},
                new User(){ Name = "jule", Email = "jule@test.fr", Phone = "776554433"},
            };

            return users;
        }

        public ActionResult ExportToExcel()
        {
            var users = GetUsers();

            var stream = new MemoryStream();

            using (var excelPackage = new ExcelPackage(stream))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("users");
                var customStyle = excelPackage.Workbook.Styles.CreateNamedStyle("customStyle");
                customStyle.Style.Font.UnderLine = true;
                customStyle.Style.Font.Color.SetColor(Color.Red);

                workSheet.Cells["A1"].Value = "sample user export";
                using (var r = workSheet.Cells["A1:C1"])
                {
                    r.Merge = true;
                    r.Style.Font.Color.SetColor(Color.Green);
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                }

                workSheet.Cells["A4"].Value = "Name";
                workSheet.Cells["B4"].Value = "Email";
                workSheet.Cells["C4"].Value = "Phone";
                workSheet.Cells["A4:C4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["A4:C4"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                int startRow = 5;
                int row = startRow;
                foreach (var user in users)
                {
                    workSheet.Cells[row, 1].Value = user.Name;
                    workSheet.Cells[row, 2].Value = user.Email;
                    workSheet.Cells[row, 3].Value = user.Phone;
                    row++;
                }

                excelPackage.Workbook.Properties.Title = "user list";
                excelPackage.Workbook.Properties.Author = "Alexei";
                excelPackage.Save();
            }

            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "users.xlsx");
        }

        [HttpGet]
        public ActionResult BatchUserUpload()
        {
            return View();
        }

        [HttpPost]
        public ActionResult BatchUserUpload(IFormFile batchUser)
        {
            if (!ModelState.IsValid) return View();

            if (batchUser?.Length == 0) return View();

            Stream stream = batchUser.OpenReadStream();
            var users = new List<User>();

            try
            {
                using(var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                    int rowCount = workSheet.Dimension.Rows;

                    for (int row = 2; row < rowCount; row++)
                    {
                        try
                        {
                            string name = workSheet.Cells[row, 1].Value?.ToString();
                            string email = workSheet.Cells[row, 2].Value?.ToString();
                            string phone = workSheet.Cells[row, 3].Value?.ToString();
                            var user = new User
                            {
                                Name = name,
                                Email = email,
                                Phone = phone,
                            };
                            users.Add(user);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }

                return View("Index", users);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return View();
        }
    }
}
