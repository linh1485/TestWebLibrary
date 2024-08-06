using OpenQA.Selenium.Edge;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;

namespace TestWebLibrary
{
    internal class test
    {
        protected string localHost = "http://localhost:4200";
        protected IWebDriver driver;
        protected string pathOfExcel;
        protected string[] newString;

        [SetUp]
        public void Setup()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string currentDirectory = Directory.GetCurrentDirectory();
            pathOfExcel = "C:\\Users\\doail\\OneDrive\\Tài liệu\\CNTT\\Bảo đảm chất lượng phần mềm\\TestWebLibrary\\21DH113829_DoAiLinh.xlsx";
            pathOfExcel = Path.Combine(currentDirectory, pathOfExcel);//đường dẫn tuyệt đối

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            ChromeDriverService service = ChromeDriverService.CreateDefaultService("C:\\Users\\doail\\OneDrive\\Tài liệu\\CNTT\\Bảo đảm chất lượng phần mềm\\chromedriver-win64\\chromedriver-win64");
            //open chrome: https://googlechromelabs.github.io/chrome-for-testing/
            //menu stable, choose chrome driver win64, download
            ChromeDriver chromeDriver = new ChromeDriver(service, options);
            driver = chromeDriver;

        }

        public string[] ConvertToArray(string[] parts)
        {
            string[] newString = new string[parts.Length];
            for (int j = 0; j < parts.Length; j++)
            {
                if (parts[j] == "null")
                {
                    newString[j] = "";
                }
                else
                {
                    newString[j] = parts[j];
                }
                Console.WriteLine(newString[j]);
            }
            return newString;
        }

        public bool CompareExpectedAndActual(string expected, string actual)
        {
            if (expected == actual) return true;
            else return false;
        }


        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}
