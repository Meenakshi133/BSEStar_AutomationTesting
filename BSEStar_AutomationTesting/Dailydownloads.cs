using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection.Metadata;
using BSEStar_AutomationTesting.Utilities;
using static System.Net.Mime.MediaTypeNames;
using OpenQA.Selenium.Firefox;

namespace BSEStar_AutomationTesting
{
    public class Dailydownloads
    {
        private WebDriver driver;
        private WebDriverWait wait;
        Random random = new Random();
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://bsestarmf.in/");
            driver.Manage().Window.Maximize();
            LoginBSEStar("1280703", "12807", "123#456");
        }

        [TearDown]
        public void Teardown()
        {
            driver.Quit();
        }

        public bool isAlertPresent(WebDriver driver, out string alertMessage)
        {
            try
            {
                IAlert alert = driver.SwitchTo().Alert();
                alertMessage = alert.Text;
                return true;
            }
            catch (NoAlertPresentException)
            {
                alertMessage = null;
                return false;
            }
          
        }

        private void LoginBSEStar(string userID, string memberID, string password)
        {
            SendKeysById("txtUserId", userID);
            SendKeysById("txtMemberId", memberID);
            SendKeysById("txtPassword", password);
            ClickById("btnLogin");
            WaitForSeconds(4);
        }

        private void SendKeysById(string elementId, string text)
        {
            driver.FindElement(By.Id(elementId)).SendKeys(text);
        }

        private void ClickById(string elementId)
        {
            driver.FindElement(By.Id(elementId)).Click();
        }
        private void ClickExportButton(string elementId)
        {
            driver.FindElement(By.Id(elementId)).Click();
        }

        private void WaitForSeconds(int seconds)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(seconds)).Until(driver => true);
        }

        private void MoveToElementAndClick(By by)
        {
            new Actions(driver).MoveToElement(driver.FindElement(by)).Click().Perform();
        }

        private void SwitchToIframeByXpath(string xpath)
        {
            driver.SwitchTo().Frame(driver.FindElement(By.XPath(xpath)));
        }

        [Test]
        public void ValidLogin()
        {
            LoginBSEStar("1280703", "12807", "123#456");
        }

        [Test]
        public void TestDailyDownloads()
        {
            NavigateToDailyDownloads("X-SIP/I-SIP Registration Report New");
            string destinationPath = HandleFileDownload("XISIPREGRPT", "XSIP", "btnExcel","xlsx");
            LoginWAdmin("shammi.patne@gmail.com", "URSecUpgrade@2024");
            UploadBSEstarFiles(destinationPath,"");
        }
        [Test]
        public void TestXSIPCancellationDownloads()
        {
            NavigateToXSIPCancellation("X-SIP/I-SIP Cancellation Report");
            string destinationPath = HandleFileDownload("XISIPCXLRPT", "XSIPCancel", "btnDownload",".txt");
            LoginWAdmin("shammi.patne@gmail.com", "URSecUpgrade@2024");
            UploadBSEstarFiles(destinationPath,"");
        }

        [Test]
        public void TestMandateDownloadUpload()
        {
            NavigateToMandate();
            string destinationPath = HandleFileDownload("Mandate", "Mandate", "btnExportToExcel", "xlsx");
            LoginWAdmin("shammi.patne@gmail.com", "URSecUpgrade@2024");
            UploadBSEstarFiles(destinationPath, "Mandate (xlsx)");
        }

        [Test]
        public void TestSTPDownloadUpload()
        {
            NavigateToSTPRegistration();
            string destinationPath = HandleFileDownload("STP", "STP", "btnExcel", "xlsx");
            if (!string.IsNullOrEmpty(destinationPath))
            {
                LoginWAdmin("shammi.patne@gmail.com", "URSecUpgrade@2024");
                UploadBSEstarFiles(destinationPath, "Mandate (xlsx)");
            }
            else
            {
                Console.WriteLine("No files downloaded");
            }
          
        }

        private void NavigateToDailyDownloads(string mainMenu)
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(driver.FindElement(By.XPath("//a[contains(.,'Daily Downloads')]"))).Perform();
            WaitForSeconds(1);
            actions.MoveToElement(driver.FindElement(By.XPath("//body/ul[@id='ddsubmenuDlyDnld']/li[7]/a[1]/img[1]"))).Perform();
            WaitForSeconds(3);
            MoveToElementAndClick(By.XPath($"//a[contains(text(),'{mainMenu}')]"));
            WaitForSeconds(5);
        }

        private void NavigateToXSIPCancellation(string mainMenu)
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(driver.FindElement(By.XPath("//a[contains(.,'Daily Downloads')]"))).Perform();
            WaitForSeconds(5);
            actions.MoveToElement(driver.FindElement(By.XPath("//body/ul[@id='ddsubmenuDlyDnld']/li[7]/a[1]/img[1]"))).Perform();
            WaitForSeconds(6);
            MoveToElementAndClick(By.XPath("(//a[contains(.,'Mandate Download')])[3]"));
           Thread.Sleep(10000);
        }

        
        private void NavigateToMandate()
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(driver.FindElement(By.XPath("//a[@href='Blank.aspx'][contains(.,'Systematic Investment')]"))).Perform();
            WaitForSeconds(5);
            actions.MoveToElement(driver.FindElement(By.XPath("(//a[contains(.,'Mandate')])[1]"))).Perform();
            WaitForSeconds(6);
            MoveToElementAndClick(By.XPath("//a[contains(.,'Mandate Search')]"));
           Thread.Sleep(10000);
        }
        
        private void NavigateToSTPRegistration()
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(driver.FindElement(By.XPath("//a[contains(.,'Daily Downloads')]"))).Perform();
            WaitForSeconds(5);
            actions.MoveToElement(driver.FindElement(By.XPath("(//a[contains(.,'STP')])[12]"))).Perform();
            WaitForSeconds(6);
            MoveToElementAndClick(By.XPath("//a[contains(.,'STP Registration Report New')]"));
           Thread.Sleep(10000);
        }


        private string HandleFileDownload(string downloadedfileName, string newFile, string clickBtnExport, string fileExtension)
        {
            Console.WriteLine(driver.Title);
            IReadOnlyList<IWebElement> iframes = driver.FindElements(By.TagName("iframe"));
            Console.WriteLine($"Number of iframes on the page: {iframes.Count}");
            SwitchToIframeByXpath("/html/body/iframe[1]");
            string currentDate = DateTime.Now.ToString("dd/MMM/yyyy");
            SetDate(By.Id("txtFromDate"), "01-May-2024");
            SetDate(By.Id("txtToDate"), "24-jun-2024");

            string downloadDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            List<FileInfo> filesBefore = Directory.GetFiles(downloadDirectory).Select(f => new FileInfo(f)).ToList();
            //driver.FindElement(By.Id("btnExcel")).Click();
            ClickExportButton(clickBtnExport);
            WaitForSeconds(5);
            string destinationPath = string.Empty;
            try
            {
                if (isAlertPresent(driver, out string alertMessage))
                {
                    WaitForSeconds(3);
                    Console.WriteLine(alertMessage);
                    Teardown();
                }
                else
                {
                    FileInfo newestFile = WaitForFileDownload(downloadDirectory, filesBefore, downloadedfileName);

                    if (newestFile != null)
                    {
                        string newFilename = $"{newFile}_{random.Next(10000000)}.{fileExtension}";
                        destinationPath = MoveFileToDestination(newestFile.FullName, newFilename);
                    }
                }
               
               
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
            }
            finally
            {

                if (driver != null)
                {
                    driver.Quit();
                    driver = null;
                }

            }
            return destinationPath;
        }


        private void SetDate(By elementId, string date)
        {
            var dates = driver.FindElements(elementId);
            if (dates.Count > 0)
            {
                IWebElement fromDateElement = dates[0];
                if (fromDateElement.Displayed)
                {
                    fromDateElement.Clear();
                    fromDateElement.SendKeys(date);
                }
                
                else
                {
                    Console.WriteLine(elementId + "is not visible");
                }

            }
        }

        private FileInfo WaitForFileDownload(string downloadDirectory, List<FileInfo> filesBefore, string fileNameContains)
        {
            for (int i = 0; i < 30; i++) // Check every second for 30 seconds
            {
                List<FileInfo> filesAfter = Directory.GetFiles(downloadDirectory).Select(f => new FileInfo(f)).ToList();
                FileInfo newestFile = filesAfter.Except(filesBefore).OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                if (newestFile != null && newestFile.FullName.Contains(fileNameContains))
                {
                    return newestFile;
                }

                Thread.Sleep(1000);
            }

            return null;
        }

        private string MoveFileToDestination(string sourceFilePath, string newFilename)
        {
            string destinationDirectory = constant.FilePath;

            string createDirectoryName = Path.Combine(destinationDirectory, DateTime.Now.ToString("yyyMMdd"));

            if (!Directory.Exists(createDirectoryName))
            {
                Directory.CreateDirectory(createDirectoryName);
            }

            string destinationPath = Path.Combine(createDirectoryName, newFilename);

            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }

            File.Move(sourceFilePath, destinationPath);
            Console.WriteLine($"File moved to {destinationPath}");
            return destinationPath;
        }


        public void LoginWAdmin(string userName, string password)

        {
            driver.Navigate().GoToUrl("https://adminuat.rsec.co.in");
            Actions actions = new Actions(driver);
            WaitForSeconds(3);
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

        public void UploadBSEstarFiles(string filePath, string fileTypeDDL)
        {
            driver.Manage().Window.Maximize();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Navigate().GoToUrl("https://adminuat.rsec.co.in/BSEStar/BSEFileUpload");
            Thread.Sleep(3000);
            SelectDropdownOptionById("select2-ddl_filetype-container", fileTypeDDL);
            SelectDropdownOptionById("select2-ddl_clienttype-container", "Executionary");
            Thread.Sleep(3000);
            IWebElement Fileupload = driver.FindElement(By.Id("file_upload"));
            Fileupload.SendKeys(filePath);

            driver.FindElement(By.Id("btn_upload")).Click();
            Thread.Sleep(3000);
        }
        public void UploadFiles(string filePath)
        {
            driver.Manage().Window.Maximize();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Navigate().GoToUrl("https://adminuat.rsec.co.in/BSEStar/BSEFileUpload");
            Thread.Sleep(3000);
            IWebElement FileType = driver.FindElement(By.Id("select2-ddl_filetype-container"));
            FileType.Click();
            string optionXpath = $"//li[contains(.,'XSIP Cancellation (txt)')]";
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
            IWebElement Fileupload = driver.FindElement(By.Id("file_upload"));
            Fileupload.SendKeys(filePath);

            driver.FindElement(By.Id("btn_upload")).Click();
            Thread.Sleep(3000);
        }
    

        public void SelectDropdownOptionById(string dropdownId, string optionText)
        {
            IWebElement fieldsDDL = driver.FindElement(By.Id(dropdownId));
            if(fieldsDDL.Displayed)
            {

                fieldsDDL.Click();
                string OptionXpath = $"//li[contains(.,'{optionText}')]";
                IWebElement optionElement = driver.FindElement(By.XPath(OptionXpath));
                optionElement?.Click();
            }
            else
            {
                Console.WriteLine("Dropdown is not present");
            }
        }
    }




}
