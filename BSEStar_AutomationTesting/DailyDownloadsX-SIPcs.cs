using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace BSEStar_AutomationTesting
{
    [TestFixture]
    public class DailyDownloadsX_SIPcs
    {


        private WebDriver driver;

        [SetUp]
        public void Setup()
        {

            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://bsestarmf.in/");
            driver.Manage().Window.Maximize();

        }

        [TearDown]
        public void Teardown()
        {
            driver.Quit();
        }


        public void LoginBSEStar(string UserID, string MemberID, string Password)
        {
            driver.FindElement(By.Id("txtUserId")).SendKeys(UserID);
            driver.FindElement(By.Id("txtMemberId")).SendKeys(MemberID);
            driver.FindElement(By.Id("txtPassword")).SendKeys(Password);
            driver.FindElement(By.Id("btnLogin")).Click();
            Thread.Sleep(4000);
        }


        public void ValidLogin()
        {
            LoginBSEStar("1280703", "12807", "123#456");
        }

        [Test]
        public void TestDailydownloads()
        {

            LoginBSEStar("1280703", "12807", "123#456");
            driver.Manage().Window.Maximize();
            Actions actions = new Actions(driver);
            IWebElement dailyDownloadsMenu = driver.FindElement(By.XPath("//a[@href='Blank.aspx'][contains(.,'Daily Downloads')]"));
            actions.MoveToElement(dailyDownloadsMenu).Perform();
            Thread.Sleep(1000);
            IWebElement xsipMenu = driver.FindElement(By.XPath("//body/ul[@id='ddsubmenuDlyDnld']/li[7]/a[1]/img[1]"));
            actions.MoveToElement(xsipMenu).Perform();
            Thread.Sleep(3000);
            IWebElement xsipRegistrationNew = driver.FindElement(By.XPath("//a[contains(text(),'X-SIP/I-SIP Registration Report New')]"));
            actions.MoveToElement(xsipRegistrationNew).Perform();
            xsipRegistrationNew.Click();
            Thread.Sleep(5000);
            Console.WriteLine(driver.Title);
            IReadOnlyList<IWebElement> iframes = driver.FindElements(By.TagName("iframe"));
            Console.WriteLine($"Number of iframes on the page: {iframes.Count}");
            IWebElement iframeElement = driver.FindElement(By.XPath("//body/iframe[1]"));
            driver.SwitchTo().Frame(iframeElement);
            string currentDate = DateTime.Now.ToString("dd/MMM/yyyy");
            IWebElement fromdate = driver.FindElement(By.Id("txtFromDate"));
            fromdate.Clear();
            fromdate.SendKeys("31-May-2024");
            Thread.Sleep(5000);
            IWebElement todate = driver.FindElement(By.Id("txtToDate"));
            todate.Clear();
            todate.SendKeys(currentDate);
            Thread.Sleep(5000);
            // File downlowd
            string downloadDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            var filesBefore = Directory.GetFiles(downloadDirectory).Select(f => new FileInfo(f)).ToList();
            driver.FindElement(By.Id("btnExcel")).Click();
            Thread.Sleep(5000);
            string filename = string.Empty;
            string filepath = Path.GetFileName(downloadDirectory);
            FileInfo newestFile = null;
            string newFilename = string.Empty;
            for (int i = 0; i < 30; i++) // Check every second for 30 seconds
            {
                var filesAfter = Directory.GetFiles(downloadDirectory).Select(f => new FileInfo(f)).ToList();
                newestFile = filesAfter.Except(filesBefore).OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                if (newestFile != null && newestFile.FullName.Contains("XISIPREGRPT"))
                {
                    break;
                }
                Thread.Sleep(1000);
            }

            if (newestFile != null)
            {
                newFilename = $"XSIP_{DateTime.Now.ToString("yyyyMMdd")}.xlsx";

                Console.WriteLine($"File downloaded to {newestFile.FullName}");
                // Move file to another directory
                string destinationDirectory = @"D:\Developement\BSEStardownloadfile";
                //string TimeStamp = DateTime.Now.ToString("yyyyMMddHHmmssff");

                string createDirectoryname = Path.Combine(destinationDirectory, DateTime.Now.ToString("yyyMMdd"));
                if (!Directory.Exists(createDirectoryname))
                {
                    Directory.CreateDirectory(createDirectoryname);
                }
                string destinationPath = Path.Combine(createDirectoryname, newFilename);
                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath);
                }
                File.Move(newestFile.FullName, destinationPath);
                // Console.WriteLine($"File downloaded to {newestFile.FullName}");
                Console.WriteLine($"File moved to {destinationPath}");
            }

        }


    }
}
