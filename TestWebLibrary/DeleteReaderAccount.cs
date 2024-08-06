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
    internal class DeleteReaderAccount : test
    {
        //Xóa 1 loại sách
        [Test]
        [TestCase("admin@gmail.com", "admin123")]

        public void testDelete1ReaderAccount(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - ReaderAccount");

            string expected = worksheet.Cell(12, 3).Value.ToString();
            //Login
            driver.Navigate().GoToUrl(localHost + "/login");
            driver.FindElement(By.Id("input-text-2")).Click();
            driver.FindElement(By.Id("input-text-2")).SendKeys(username);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("input-text-3")).Click();
            driver.FindElement(By.Id("input-text-3")).SendKeys(password);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn")).Click();
            Thread.Sleep(2000);

            //ReaderAccount
            driver.FindElement(By.CssSelector(".bi")).Click();
            driver.FindElement(By.LinkText("Reader Accounts")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//tbody/tr[1]/td[4]/button[2]")).Click();
            driver.FindElement(By.XPath("//button[@class='p-element p-ripple ng-tns-c3410224651-1 p-confirm-dialog-accept p-button p-component ng-star-inserted']")).Click();
            Thread.Sleep(1000);

            if (driver.FindElement(By.XPath("//div[@class='p-toast-summary ng-tns-c229242374-15']")).Text == "Success")
            {
                string actual = "Hệ thống xóa tài khoản người đọc thành công và trả về trang Index";
                worksheet.Cell(12, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(12, 5).Value = "Passed";
                else worksheet.Cell(12, 5).Value = "Failed";
            }
            else
            {
                string actual = "Hệ thống báo lỗi không xóa được tài khoản người đọc";
                worksheet.Cell(12, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(12, 5).Value = "Passed";
                else worksheet.Cell(12, 5).Value = "Failed";
            }

            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();


        }


        [Test]
        [TestCase("admin@gmail.com", "admin123")]

        public void testDeleteAllReaderAccount(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - ReaderAccount");

            string expected = worksheet.Cell(13, 3).Value.ToString();
            driver.Navigate().GoToUrl(localHost + "/login");
            driver.FindElement(By.Id("input-text-2")).Click();
            driver.FindElement(By.Id("input-text-2")).SendKeys(username);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("input-text-3")).Click();
            driver.FindElement(By.Id("input-text-3")).SendKeys(password);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn")).Click();
            Thread.Sleep(2000);

            //ReaderAccount
            driver.FindElement(By.CssSelector(".bi")).Click();
            driver.FindElement(By.LinkText("Reader Accounts")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//button[@class='p-element p-button-secondary mx-2 p-button p-component p-button-icon-only']")).Click();
            Thread.Sleep(1000);

            if (driver.FindElement(By.XPath("//div[@class='p-toast-summary ng-tns-c229242374-15']")).Text == "Success")
            {
                string actual = "Hệ thống xóa tất cả tài khoản người đọc thành công và trả về trang Index";
                worksheet.Cell(13, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(13, 5).Value = "Passed";
                else worksheet.Cell(13, 5).Value = "Failed";
            }
            else
            {
                string actual = "Hệ thống báo lỗi không xóa được tất cả tài khoản người đọc";
                worksheet.Cell(13, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(13, 5).Value = "Passed";
                else worksheet.Cell(13, 5).Value = "Failed";
            }

            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();


        }

    }
}
