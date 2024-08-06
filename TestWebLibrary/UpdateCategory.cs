using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Bibliography;
using Bytescout.Spreadsheet;
using System.Net;

namespace TestWebLibrary
{
    internal class UpdateCategory : test
    {
        [Test]
        [TestCase("admin@gmail.com", "admin123")]

        public void testUpdateCategory(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - Category");

            for (int i = 8; i <= 10; i++)
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

                driver.FindElement(By.XPath("//tbody/tr[2]/td[4]/button[1]]")).Click();
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("//input[@id='input-text-5']")).Click();
                driver.FindElement(By.XPath("//input[@id='input-text-5']")).SendKeys(newString[0]);
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("//input[@id='input-text-6']")).Click();
                driver.FindElement(By.XPath("//input[@id='input-text-6']")).SendKeys(newString[1]);
                Thread.Sleep(1000);

                driver.FindElement(By.CssSelector(".btn")).Click();
                Thread.Sleep(2000);

                if (driver.FindElement(By.XPath("//div[@class='p-toast-summary ng-tns-c229242374-7']")).Text == "Success")
                {
                    string actual = "Hệ thống cập nhật loại sách thành công và trả về trang Index";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                else if (driver.FindElement(By.XPath("//div[@class='p-toast-summary ng-tns-c229242374-7']")).Text == "Error")
                {
                    string actual = "Hệ thống báo lỗi sai dữ liệu để cập nhật loại sách";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }

                else if (driver.Url.Contains(localHost + "/category(modal: category / edit /)"))
                {
                    string actual = "Hệ thống báo lỗi không đủ dữ liệu để cập nhật loại sách ";
                    worksheet.Cell(i, 4).Value = actual;
                    if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(i, 5).Value = "Passed";
                    else worksheet.Cell(i, 5).Value = "Failed";
                }
                else
                {
                    string actual = "Hệ thống báo lỗi sai dữ liệu để cập nhật loại sách ";
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
