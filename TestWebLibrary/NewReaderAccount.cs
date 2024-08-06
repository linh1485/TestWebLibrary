using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GSF.IO;
using DocumentFormat.OpenXml.Bibliography;
using Bytescout.Spreadsheet;

namespace TestWebLibrary
{
    internal class NewReaderAccount : test
    {
        [Test]
        [TestCase("admin@gmail.com", "admin123")]

        public void testNewReaderAccount(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - ReaderAccount");

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

                driver.FindElement(By.CssSelector(".p-button-label")).Click();
                Thread.Sleep(1000);

                driver.FindElement(By.CssSelector("div[aria-label='dropdown trigger']")).Click();
                driver.FindElement(By.CssSelector("li[aria-label='Đỗ Ái Linh']")).Click();
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("//input[@id='input-text-7']")).Click();
                driver.FindElement(By.XPath("//input[@id='input-text-7']")).SendKeys(newString[0]);
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("//input[@id='input-text-8']")).Click();
                driver.FindElement(By.XPath("//input[@id='input-text-8']")).SendKeys(newString[1]);
                Thread.Sleep(1000);

                driver.FindElement(By.CssSelector(".btn")).Click();
                Thread.Sleep(1000);

                if (driver.FindElement(By.CssSelector(".p-toast-summary.ng-tns-c229242374-13")).Text == "Success")
                {
                    string actual = "Hệ thống thêm tài khoản người đọc mới thành công và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                else if (driver.FindElement(By.CssSelector(".p-toast-summary.ng-tns-c229242374-13")).Text == "Error")
                {
                    string actual = "Hệ thống thêm tài khoản người đọc mới thành công và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                else if (driver.Url.Contains(localHost + "/reader-account-list(modal:reader-account-list/edit/)"))
                {
                    string actual = "Hệ thống báo lỗi không đủ dữ liệu để thêm tài khoản người đọc mới";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                else 
                {
                    string actual = "Hệ thống báo lỗi sai dữ liệu để thêm tài khoản người đọc mới";
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
