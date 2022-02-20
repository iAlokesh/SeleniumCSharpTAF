using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using RelevantCodes.ExtentReports;
using System.Configuration;

namespace SeleniumTestAutomationFramework.ResuableLibraries
{
    public class ReusableMethods
    {
        public static IWebDriver sDriver;

        public static IWebDriver LaunchBrowser(string BrowserName, ExtentTest Test)
        {
            if (BrowserName.Equals("chrome"))
            {
                sDriver = new ChromeDriver();
                sDriver.Manage().Window.Maximize();
                sDriver.Navigate().GoToUrl(ConfigurationManager.AppSettings["URL"]);
                Test.Log(LogStatus.Info, "Browser has been launched successfully");
                Test.Log(LogStatus.Info, Test.AddBase64ScreenCapture(GetScreenshot()));
            }
            else if (BrowserName.Equals("firefox"))
            {
                sDriver = new FirefoxDriver();
                sDriver.Manage().Window.Maximize();
                sDriver.Navigate().GoToUrl(ConfigurationManager.AppSettings["URL"]);
                Test.Log(LogStatus.Info, "Browser has been launched successfully");
                Test.Log(LogStatus.Info, Test.AddBase64ScreenCapture(GetScreenshot()));
            }
            return sDriver;
        }

        public static string GetScreenshot()
        {
            string Screenshot = ((ITakesScreenshot)sDriver).GetScreenshot().AsBase64EncodedString;
            return "data:image/jpg;base64, " + Screenshot;
        }
    }
}
