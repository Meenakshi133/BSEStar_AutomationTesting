using OpenQA.Selenium;
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
    public class UploadBSEFiles
    {
        private WebDriver driver;
        private WebDriverWait wait;

    [SetUp]
    public void Setup()
    {
        driver = new FirefoxDriver();
        driver.Navigate().GoToUrl("https://adminuat.rsec.co.in");
        wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
    }

    [TearDown]
    public void Teardown()
    {
        driver.Quit();
    }

        //Login
    private void LoginWAdmin(string userName, string password)
    {
        Actions actions = new Actions(driver);
        SendKeysById("txtUserName", userName);
        WaitForSeconds(5);
        SendKeysById("txtPassword", password);
        WaitForSeconds(5);
        ClickById("btn_login");
    }

    private void SendKeysById(string elementId, string text)
    {
        IWebElement element = driver.FindElement(By.Id(elementId));
        element.SendKeys(text);
        //sExecutor.ExecuteScript("var event = new Event('change', { bubbles: true }); arguments[0].dispatchEvent(event);", element);
          ((IJavaScriptExecutor)driver).ExecuteScript("$(arguments[0].change());", element);
           //TriggerChangeEvent(element);
        }

        private void TriggerChangeEvent(IWebElement element)
        {
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
           // string script = @"
          //  jsExecutor.ExecuteScript(script, element);
        }
        private void ClickById(string elementId)
    {
        driver.FindElement(By.Id(elementId)).Click();
    }

    private void WaitForSeconds(int seconds)
    {
        new WebDriverWait(driver, TimeSpan.FromSeconds(seconds)).Until(d => true);
    }

    public void UploadBSEstarFiles()
    {
        LoginWAdmin("shammi.patne@gmail.com", "Zoi@1234");
        driver.Manage().Window.Maximize();
        driver.Navigate().GoToUrl("https://adminuat.rsec.co.in/BSEStar/BSEFileUpload");
        WaitForSeconds(3);
        Thread.Sleep(5000);

        SelectDropdownOptionById("select2-ddl_filetype-container", "XSIP Registration (xlsx)");
        WaitForSeconds(5);
        SelectDropdownOptionById("select2-ddl_clienttype-container", "Executionary");

        string filePath = GetXSIPFile();
        if (filePath != null)
        {
            IWebElement fileUploadElement = driver.FindElement(By.Id("file_upload"));
            fileUploadElement.SendKeys(filePath);
            ClickById("btn_upload");
            WaitForSeconds(3);
            
        }
        else
        {
                //TODO - We will store this in log file - Notification must send - By Tony
            Console.WriteLine("No Matching records found");
        }
    }

    private void SelectDropdownOptionById(string dropdownId, string optionText)
    {
        IWebElement dropdown = driver.FindElement(By.Id(dropdownId));
        dropdown.Click();
        string optionXpath = $"//li[contains(.,'{optionText}')]";
        IWebElement optionElement = wait.Until(d => d.FindElement(By.XPath(optionXpath)));
        optionElement.Click();
    }

    private string GetXSIPFile()
    {
        string directoryPath = @"D:\Developement\BSEStardownloadfile";
        string todayFolder = Path.Combine(directoryPath, DateTime.Now.ToString("yyyyMMdd"));

        if (Directory.Exists(todayFolder))
        {
            var directoryInfo = new DirectoryInfo(todayFolder);
            var xsipFiles = directoryInfo.GetFiles("*.xlsx").OrderByDescending(f => f.LastWriteTime);
            foreach (var file in xsipFiles)
            {
                if (file.Name.Contains("XSIP"))
                {
                    return file.FullName;
                }
            }
        }

        return null;
    }

    }
}
