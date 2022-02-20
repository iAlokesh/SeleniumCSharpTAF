using NUnit.Framework;
using SeleniumTestAutomationFramework.Testcases;

namespace SeleniumTestAutomationFramework
{
    [TestFixture]
    public class TestScripts : BaseClass
    {
        [Test]
        public void TestDemo()
        {
            Gsp.EnterValueAndSearch(TDP.GetCellData("TestDemo", "Key", 3));
        }
    }
}
