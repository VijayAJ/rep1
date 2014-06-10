using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using NUnit.Framework;
using NUnit.Core;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Configuration;
using Microsoft.CSharp.RuntimeBinder;
using System.Data.Sql;
using System.IO;
using System.Drawing.Imaging;
using Selenium;

namespace Test
{
    /// <summary>
    /// Summary description for DDTest
    /// </summary>
   //[CodedUITest]
    [TestFixture()]
    public class DDTest
    {
        public DDTest()
        {
        }

        public IWebDriver driver;
        Library.GeneralLibrary Lib = new Library.GeneralLibrary();
        private string baseURL;

        //[TestInitialize]
        [SetUp]
        public void Initilize()
        {
            Lib.driver = driver;
            //driver = new FirefoxDriver();
            driver = new InternetExplorerDriver();
            baseURL = "http://datadriver.dev.gs1us.org/";
            WebDriverWait wiat = new WebDriverWait(driver, TimeSpan.FromSeconds(300));
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(30));

        }

        #region Datadriver Testmethods
        //[TestMethod]
        [Test]
        public void TheDDScenario1Test()
        {
            //Get path of the Excel file stored in the project
            string projLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string[] projArray = Regex.Split(projLocation, "bin");
            //MessageBox.Show(projArray[0]);
            string filepath = projArray[0] + "Resource\\DD_test_data.xls";
           // string filepath="D:\\Selenium_test_data\\DD_test_data.xls";
            string username = Lib.Readdata(filepath, 1, 1, 2);

            //+ "/dd2/auth/Login.action"
            driver.Navigate().GoToUrl(baseURL);

            if (Lib.isElementPresent(By.Name("j_username"))==false)
            {
                driver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
            }
            //driver.SwitchTo().Window("");
            Thread.Sleep(4000);
            driver.FindElement(By.Name("j_username")).Clear();
            driver.FindElement(By.Name("j_username")).SendKeys(username);
            driver.FindElement(By.Name("j_password")).Clear();
            driver.FindElement(By.Name("j_password")).SendKeys("DEVVJ");
            driver.FindElement(By.CssSelector("input.button_default")).Click();
            Lib.WaitunitleElementPresent(By.Id("PrefixSelection_0"));
            driver.FindElement(By.Id("PrefixSelection_0")).Click();
            Lib.WaitunitleElementPresent(By.Id("MainMenu_0"));
            driver.FindElement(By.Id("MainMenu_0")).Click();
            Lib.WaitunitleElementPresent(By.Id("toggleSearchText"));
            driver.FindElement(By.Id("toggleSearchText")).Click();
            driver.FindElement(By.Id("Products_searchBrandName")).Clear();
            driver.FindElement(By.Id("Products_searchBrandName")).SendKeys("Egg");
            driver.FindElement(By.Id("Products_searchPackageEach")).Click();
            driver.FindElement(By.Id("search")).Click();
            Lib.WaitunitleElementPresent(By.XPath("//table[@class='tablebackground']"));
            string ProductName= driver.FindElement(By.XPath("//table[@class='tablebackground']/tbody/tr[2]/td[4]")).Text.ToString();
            if (ProductName == "Egg")
            {
                Lib.Logs(TestContext.TestName +": "+"Product is searched succesfuly", "PASSED");
            }
            else
            {
                Lib.Logs(TestContext.TestName +": "+"Product is searched succesfuly", "FAILED");
            }
        }

        #endregion

        //[TestCleanup]
        [TearDown]
        public void TeardownTest()
        {
            try
            {
                driver.Quit();
            }
            catch (Exception)
            {
                // Ignore errors if unable to close the browser
            }
        }


        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;
    }
}
