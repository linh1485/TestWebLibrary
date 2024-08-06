using DocumentFormat.OpenXml.Bibliography;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bytescout.Spreadsheet;

namespace TestWebLibrary
{
    internal class SearchReaderAccount : test
    {
        [Test]
        [TestCase("admin@gmail.com", "admin123")]
        public void testSearchReaderAccountName(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - SearchReaderAccount");

            for (int i = 1; i <= 6; i++)
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

                driver.FindElement(By.CssSelector(".bi")).Click();
                driver.FindElement(By.LinkText("Reader Accounts")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-3.ng-star-inserted")).Click();
                driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-3.ng-star-inserted")).Click();

                driver.FindElement(By.CssSelector("p-columnfilterformelement[class='p-element p-fluid ng-tns-c2058025319-3 ng-star-inserted'] input[type='text']")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("p-columnfilterformelement[class='p-element p-fluid ng-tns-c2058025319-3 ng-star-inserted'] input[type='text']")).SendKeys(newString[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".p-column-filter-menu-button-active")).Click();
                Thread.Sleep(1000);

                string filterMatchMode = ".p-highlight";
                driver.FindElement(By.CssSelector($"{filterMatchMode}")).Click();
                Thread.Sleep(1000);
                string searchResult = driver.FindElement(By.CssSelector(".ml-1.text-global.fw-bold.custom-cursor-on-hover")).Text;
                bool result = true;


                switch (filterMatchMode)
                {
                    case ".p-highlight":
                        result = searchResult.StartsWith(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(2)":
                        result = searchResult.Contains(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(3)":
                        result = !searchResult.Contains(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(4)":
                        result = !searchResult.EndsWith(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(5)":
                        result = searchResult.Equals(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;
                    case ".p-column-filter-row-item:nth-child(6)":
                        result = !searchResult.Equals(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;
                }

                if (result)
                {
                    string actual = "Hệ thống tìm kiếm tài khoản người đọc thành công";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống báo lỗi không tìm kiếm được tài khoản người đọc";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();


            }
        }

        [Test]
        [TestCase("admin@gmail.com", "admin123")]
        public void testSearchReaderAccountID(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - SearchReaderAccount");

            for (int i = 7; i <= 12; i++)
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

                driver.FindElement(By.CssSelector(".bi")).Click();
                driver.FindElement(By.LinkText("Reader Accounts")).Click();
                Thread.Sleep(1000);
                //driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-3.ng-star-inserted")).Click();
                //driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-3.ng-star-inserted")).Click();

                driver.FindElement(By.CssSelector("p-columnfilterformelement[class='p-element p-fluid ng-tns-c2058025319-4 ng-star-inserted'] input[type='text']")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("p-columnfilterformelement[class='p-element p-fluid ng-tns-c2058025319-4 ng-star-inserted'] input[type='text']")).SendKeys(newString[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-4.ng-star-inserted")).Click();
                Thread.Sleep(1000);

                string filterMatchMode = ".p-highlight";
                driver.FindElement(By.CssSelector($"{filterMatchMode}")).Click();
                Thread.Sleep(1000);
                string searchResult = driver.FindElement(By.CssSelector("td:nth-child(2)")).Text;
                bool result = true;


                switch (filterMatchMode)
                {
                    case ".p-highlight":
                        result = searchResult.StartsWith(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(2)":
                        result = searchResult.Contains(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(3)":
                        result = !searchResult.Contains(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(4)":
                        result = !searchResult.EndsWith(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(5)":
                        result = searchResult.Equals(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;
                    case ".p-column-filter-row-item:nth-child(6)":
                        result = !searchResult.Equals(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;
                }

                if (result)
                {
                    string actual = "Hệ thống tìm kiếm tài khoản người đọc thành công";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống báo lỗi không tìm kiếm được tài khoản người đọc";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                // Save document
                spreadsheet.SaveAs(pathOfExcel);
                spreadsheet.Close();


            }
        }

        [Test]
        [TestCase("admin@gmail.com", "admin123")]
        public void testSearchReaderAccountEmail(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - SearchReaderAccount");
            int worksheetCount = worksheet.UsedRangeRowMax;
            Console.WriteLine(worksheetCount);

            for (int i = 13; i <= worksheetCount; i++)
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

                driver.FindElement(By.CssSelector(".bi")).Click();
                driver.FindElement(By.LinkText("Reader Accounts")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-3.ng-star-inserted")).Click();
                driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-3.ng-star-inserted")).Click();

                driver.FindElement(By.CssSelector("p-columnfilterformelement[class='p-element p-fluid ng-tns-c2058025319-5 ng-star-inserted'] input[type='text']")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("p-columnfilterformelement[class='p-element p-fluid ng-tns-c2058025319-5 ng-star-inserted'] input[type='text']")).SendKeys(newString[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".p-column-filter-menu-button.p-link.ng-tns-c2058025319-5.ng-star-inserted")).Click();
                Thread.Sleep(1000);

                string filterMatchMode = ".p-highlight";
                driver.FindElement(By.CssSelector($"{filterMatchMode}")).Click();
                Thread.Sleep(1000);
                string searchResult = driver.FindElement(By.CssSelector("td:nth-child(2)")).Text;
                bool result = true;


                switch (filterMatchMode)
                {
                    case ".p-highlight":
                        result = searchResult.StartsWith(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(2)":
                        result = searchResult.Contains(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(3)":
                        result = !searchResult.Contains(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(4)":
                        result = !searchResult.EndsWith(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;

                    case ".p-column-filter-row-item:nth-child(5)":
                        result = searchResult.Equals(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;
                    case ".p-column-filter-row-item:nth-child(6)":
                        result = !searchResult.Equals(newString[0], StringComparison.OrdinalIgnoreCase); // Không phân biệt hoa thường
                        break;
                }

                if (result)
                {
                    string actual = "Hệ thống tìm kiếm tài khoản người đọc thành công";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống báo lỗi không tìm kiếm được tài khoản người đọc";
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
