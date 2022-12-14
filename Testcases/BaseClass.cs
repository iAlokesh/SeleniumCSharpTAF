using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumTestAutomationFramework.PageObjects;
using SeleniumTestAutomationFramework.ResuableLibraries;
using RelevantCodes.ExtentReports;
using System.Configuration;
using System.IO;
using System;
using NUnit.Framework.Interfaces;

namespace SeleniumTestAutomationFramework.Testcases
{
    public class BaseClass
    {
        public GoogleSearchPage Gsp;
        public IWebDriver Driver;
        public ExtentReports Extent;
        public ExtentTest Test;
        public TestDataProvider TDP;

        [OneTimeSetUp]
        public void ReportInitialize()
        {
            string projectDirectory = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName;
            Console.WriteLine("Project Directory: " + projectDirectory);
            Extent = new ExtentReports("D:\\CSharp\\SeleniumTestAutomationFramework\\TestReports\\Google.html", true);
        }

        [SetUp]
        public void SetUp()
        {
            Test = Extent.StartTest(TestContext.CurrentContext.Test.Name);
            Driver = ReusableMethods.LaunchBrowser(ConfigurationManager.AppSettings["BrowserName"], Test);
            TDP = new TestDataProvider("D:\\CSharp\\SeleniumTestAutomationFramework\\TestDataSheet\\TestData.xlsx");
            Gsp = new GoogleSearchPage(Driver, Test);
        }

        [TearDown]
        public void QuitBrowser()
        {
            if (TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Passed)
            {
                Test.Log(LogStatus.Pass, TestContext.CurrentContext.Test.Name + " has been passed successfully");
                Test.Log(LogStatus.Info, Test.AddBase64ScreenCapture(ReusableMethods.GetScreenshot()));
            }
            else if (TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Failed)
            {
                Test.Log(LogStatus.Fail, TestContext.CurrentContext.Test.Name + " has been failed " + TestContext.CurrentContext.Result.Message);
                Test.Log(LogStatus.Info, Test.AddBase64ScreenCapture(ReusableMethods.GetScreenshot()));
            }
            Extent.EndTest(Test);
            Driver.Quit();
        }

        [OneTimeTearDown]
        public void EndReport()
        {
            Extent.Flush();
            Extent.Close();
        }
    }
}
