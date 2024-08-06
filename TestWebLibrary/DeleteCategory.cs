using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bytescout.Spreadsheet;

namespace TestWebLibrary
{
    internal class DeleteCategory : test
    {
        //Xóa 1 loại sách
        [Test]
        [TestCase("admin@gmail.com", "admin123")]

        public void testDelete1Category(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - Category");

            string expected = worksheet.Cell(11, 3).Value.ToString();
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

            driver.FindElement(By.XPath("//tbody/tr[2]/td[4]/button[2]")).Click();
            driver.FindElement(By.XPath("//button[@class='p-element p-ripple ng-tns-c3410224651-1 p-confirm-dialog-accept p-button p-component ng-star-inserted']")).Click();
            Thread.Sleep(1000);

            if (driver.FindElement(By.XPath("//div[@class='p-toast-summary ng-tns-c229242374-15']")).Text == "Success")
            {
                string actual = "Hệ thống xóa loại sách thành công và trả về trang Index";
                worksheet.Cell(11, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(11, 5).Value = "Passed";
                else worksheet.Cell(11, 5).Value = "Failed";
            }
            else
            {
                string actual = "Hệ thống báo lỗi không xóa được loại sách";
                worksheet.Cell(11, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(11, 5).Value = "Passed";
                else worksheet.Cell(11, 5).Value = "Failed";
            }

            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();
        }

        //Xóa tất cả loại sách
        [Test]
        [TestCase("admin@gmail.com", "admin123")]

        public void testDeleteAllCategory(string username, string password)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.LoadFromFile(@$"{pathOfExcel}");
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("Ai Linh - Category");

            string expected = worksheet.Cell(12, 3).Value.ToString();
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

            driver.FindElement(By.XPath("//div[@class='p-checkbox-box custom-cursor-on-hover']")).Click();
            Thread.Sleep(1000);

            driver.FindElement(By.XPath("//button[@class='p-element p-button-secondary mx-2 p-button p-component p-button-icon-only ng-star-inserted']")).Click();
            Thread.Sleep(1000);

            if (driver.FindElement(By.XPath("//div[@class='p-toast-summary ng-tns-c229242374-15']")).Text == "Success")
            {
                string actual = "Hệ thống xóa tất cả loại sách thành công và trả về trang Index";
                worksheet.Cell(12, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(12, 5).Value = "Passed";
                else worksheet.Cell(12, 5).Value = "Failed";
            }
            else
            {
                string actual = "Hệ thống báo lỗi không xóa được tất cả loại sách";
                worksheet.Cell(12, 4).Value = actual;
                if (CompareExpectedAndActual(expected, actual)) worksheet.Cell(12, 5).Value = "Passed";
                else worksheet.Cell(12, 5).Value = "Failed";
            }

            // Save document
            spreadsheet.SaveAs(pathOfExcel);
            spreadsheet.Close();


        }

    }
}
