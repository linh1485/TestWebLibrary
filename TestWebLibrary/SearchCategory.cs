using DocumentFormat.OpenXml.Bibliography;
using GSF.IO;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bytescout.Spreadsheet;


namespace TestWebLibrary
{
    internal class SearchCategory : test
    {
        [Test]
        [TestCase("admin@gmail.com", "admin123")]
        public void testSearchCategory(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Bytescout.Spreadsheet.Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - SearchCategory");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);

            for (int i = 1; i <= worksheetCount; i++)
            {
                string expected = worksheet.Cell(i, 3).Value.ToString();
                string cellValues = worksheet.Cell(i, 2).Value.ToString();
                string[] parts = cellValues.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                string[] newString = ConvertToArray(parts);
                driver.Navigate().GoToUrl(localHost + "/login");
                driver.FindElement(By.Id("input-text-2")).Click();
                driver.FindElement(By.Id("input-text-2")).SendKeys(username);
                Thread.Sleep(1000);
                driver.FindElement(By.Id("input-text-3")).Click();
                driver.FindElement(By.Id("input-text-3")).SendKeys(password);
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".btn")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.CssSelector("i[class='fa fa-book']")).Click();
                Thread.Sleep(1000);

                driver.FindElement(By.CssSelector(".icon-arrow-container > .fa")).Click();
                driver.FindElement(By.LinkText("Categories")).Click();
                Thread.Sleep(1000);

                driver.FindElement(By.CssSelector("input[placeholder='Search...']")).Click();
                driver.FindElement(By.CssSelector("input[placeholder='Search...']")).SendKeys(newString[0]);


                if (driver.FindElement(By.CssSelector("tbody tr:nth-child(1) td:nth-child(2)")).Text == newString[0])
                {
                    string actual = "Hệ thống tìm kiếm loại sách thành công";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống báo lỗi không tìm kiếm được loại sách mới";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();


            }
        }


    }
}
