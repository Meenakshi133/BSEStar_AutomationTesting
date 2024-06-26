using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSEStar_AutomationTesting
{
    [TestFixture]
    public class WadminBSEUpload
    {
        private WebDriver driver;

        [SetUp]
        public void Setup()
        {
            driver = new FirefoxDriver();
            //driver.Navigate().GoToUrl("https://adminchola.jademoney.in");
            driver.Navigate().GoToUrl("https://adminuat.rsec.co.in");
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }



        public void LoginWAdmin(string userName, string password)

        {
            Actions actions = new Actions(driver);
            IWebElement user = driver.FindElement(By.Id("txtUserName"));
            user.SendKeys(userName);
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript("$(arguments[0]).change();", user);
            // actions.MoveToElement(user, 1, 1).Click().Perform();
            Thread.Sleep(5000);
            //actions.MoveToElement(user).Perform();
            // driver.FindElement(By.Id("txtUserName")).
            IWebElement pswd = driver.FindElement(By.Id("txtPassword"));
            pswd.SendKeys(password);
            jsExecutor.ExecuteScript("$(arguments[0]).change();", pswd);
            //actions.MoveToElement(pswd).Perform();
            Thread.Sleep(5000);
            driver.FindElement(By.Id("btn_login")).Click();

        }

        [Test]
        public void UploadBSEstarFiles()
        {
            LoginWAdmin("shammi.patne@gmail.com", "URSecUpgrade@2024");
            driver.Manage().Window.Maximize();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Navigate().GoToUrl("https://adminuat.rsec.co.in/BSEStar/BSEFileUpload");
            Thread.Sleep(3000);
            IWebElement FileType = driver.FindElement(By.Id("select2-ddl_filetype-container"));
            FileType.Click();
            string optionXpath = "//li[contains(.,'XSIP Cancellation (txt)')]";
            IWebElement? FileTypeDDL = wait.Until(d =>
            {
                IWebElement element = d.FindElement(By.XPath(optionXpath));
                if (element.Displayed && element.Enabled)
                {
                    return element;
                }
                return null;
            });
            string selectedFileType = FileTypeDDL.Text.Trim();
            FileTypeDDL?.Click();
            Thread.Sleep(3000);
            //IWebElement ClientType = driver.FindElement(By.Id("select2-ddl_clienttype-container"));
            //ClientType.Click();
            //string clientOptionXpath = "//li[contains(.,'Executionary')]";
            ////string clientOptionXpath = "//li[contains(.,'Advisory')]";
            //IWebElement? ClientTypeDDL = wait.Until(d =>
            //{
            //    IWebElement elementddl = d.FindElement(By.XPath(clientOptionXpath));
            //    if (elementddl.Displayed && elementddl.Enabled)
            //    {
            //        return elementddl;
            //    }
            //    return null;
            //});
            //ClientTypeDDL?.Click();
            Thread.Sleep(3000);
            IWebElement Fileupload = driver.FindElement(By.Id("file_upload"));
            string directoryPath = @"D:\Developement\BSEStardownloadfile";
            if (selectedFileType != null)
            {
                string todayFolder = Path.Combine(directoryPath, DateTime.Now.ToString("yyyyMMdd"));
                if (!string.IsNullOrEmpty(todayFolder))
                {
                    string getFile = string.Empty;
                    var directoryInfo = new DirectoryInfo(todayFolder);
                    var xsipFiles = directoryInfo.GetFiles("*.xlsx").OrderByDescending(f => f.LastWriteTime);
                    foreach (var files in xsipFiles)
                    {
                        if (files.Name.Contains("XSIP"))
                        {
                            getFile = Path.Combine(todayFolder, files.Name);
                            Fileupload.SendKeys(getFile);
                            break;
                        }
                        
                    }
                    if (string.IsNullOrEmpty(getFile))
                    {
                        Console.WriteLine("No Matching records found");
                    }

                }

            }


            driver.FindElement(By.Id("btn_upload")).Click();
            Thread.Sleep(3000);
            
        }
    }
}
