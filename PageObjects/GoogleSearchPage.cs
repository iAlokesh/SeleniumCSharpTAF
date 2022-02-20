using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using RelevantCodes.ExtentReports;
using SeleniumTestAutomationFramework.ResuableLibraries;

namespace SeleniumTestAutomationFramework.PageObjects
{
    public class GoogleSearchPage
    {
        public IWebDriver Driver;
        public ExtentTest Test;
        public GoogleSearchPage(IWebDriver Driver, ExtentTest Test)
        {
            this.Driver = Driver;
            this.Test = Test;
            PageFactory.InitElements(Driver, this);
        }

        [FindsBy(How = How.Name, Using = "q")]
        [CacheLookup]
        private IWebElement SearchBar { get; set; }

        public void EnterValueAndSearch(string key)
        {
            SearchBar.Click();
            SearchBar.SendKeys(key);
            SearchBar.SendKeys(Keys.Enter);
            Test.Log(LogStatus.Info, "Google Search");
            Test.Log(LogStatus.Info, Test.AddBase64ScreenCapture(ReusableMethods.GetScreenshot()));
        }
    }
}
