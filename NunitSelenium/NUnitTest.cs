using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Excel = Microsoft.Office.Interop.Excel;
using worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using workbook = Microsoft.Office.Interop.Excel.Workbook;
using System.Runtime.InteropServices;

namespace NunitSelenium 
{
    [TestFixture]
    class NUnitTest
    {
        IWebDriver driver;
        IWebElement Element;
        IAlert Alert;        
        StringBuilder verificationErrors;
        bool acceptNextAlert = true;
        AllMethods Method = new AllMethods();

        [SetUp]
        public void Initialize()
        {
            driver = new ChromeDriver();
            verificationErrors = new StringBuilder();
        }

        [Test]
        public void NameSelector()
        {
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/selectors/name/");
                //driver.Manage().Window.Maximize();
                Thread.Sleep(1000);
                Element = driver.FindElement(By.Name("myName"));
                if (Element.Displayed)
                {        
                    Method.GreenMessage("Element Located");
                }
                else
                {
                    Method.RedMessage("Element Not Located");
                }
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }           
        }        
        
        [Test]
        public void ReadWriteExcel_Selector()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            int RowCount, ColumnCount;
            String Url;

            try
            {                
                xlApp = new Excel.Application();

                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\ramya.prabhakara\Downloads\C#Programs\NunitSelenium\NunitSelenium\ExcelFiles\TestDoc.xlsx");
                //xlWorkSheet = xlWorkBook.Sheets[2];
                xlWorkSheet = xlWorkBook.Worksheets.get_Item("UrlTestSheet");
                range = xlWorkSheet.UsedRange;
                for (RowCount = 1; RowCount <= range.Rows.Count; RowCount ++)
                {
                    for (ColumnCount = 1; ColumnCount <= range.Columns.Count; ColumnCount ++)
                    {
                        Url = (range.Cells[RowCount, ColumnCount] as Excel.Range).Value;                        
                        driver.Navigate().GoToUrl(Url);

                        //Explicit Wait
                        /*WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementExists((By.Id("abc"))));
                         or
                        wait.Until(driver => driver.FindElement(By.ClassName("abc"))); */


                        //Implicit Wait
                        //driver.Manage().Timeouts().ImplicitWait(TimeSpan.FromSeconds(20));
                         
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);

                        range.Cells[RowCount, ColumnCount + 1] = "PASS";
                        Console.WriteLine("Navigated to " + Url);
                    }
                }
                xlWorkBook.Save();
                xlWorkBook.Close();                                            
                xlApp.Quit();
                if (range != null)
                    Marshal.ReleaseComObject(range);
                if (xlWorkSheet != null)
                    Marshal.ReleaseComObject(xlWorkSheet);              
                if (xlWorkBook != null)
                    Marshal.ReleaseComObject(xlWorkBook);                
                if (xlApp != null)
                    Marshal.ReleaseComObject(xlApp);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }            
        }

        [Test]
        public void IDselector()
        {
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/selectors/id/");               
                Thread.Sleep(1000);
                Element = driver.FindElement(By.Id("testImage"));
                if (Element.Displayed)
                {
                    Method.GreenMessage("Element Located");
                }
                else
                {
                    Method.RedMessage("Element Not Located");
                }
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        [Test]
        public void ClassNameSelector()
        {
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/selectors/class-name/");
                Thread.Sleep(1000);
                Element = driver.FindElement(By.ClassName("testClass"));
                if (Element.Displayed)
                {
                    Method.GreenMessage("ClassName Located");
                    Method.GreenMessage(Element.Text);
                }
                else
                {
                    Method.RedMessage("ClassName Not Located");
                }
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        [Test]
        public void CSSandXPathSelector()
        {
            IWebElement CSSPath, XPath;
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/selectors/css-path/");
                Thread.Sleep(1000);
                XPath = driver.FindElement(By.XPath("//*[@id=\"post-108\"]/div/figure/img"));
                CSSPath = driver.FindElement(By.CssSelector("#post-108 > div > figure > img"));
                if ((CSSPath.Displayed) && (XPath.Displayed))
                {
                    Method.GreenMessage("CssPath Located and XPath Located");                    
                }
                else
                {
                    Method.RedMessage("CSSPath and Xpath Not Located");
                }
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        [Test]
        public void InputTextBox()
        {            
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/selectors/name/");
                Thread.Sleep(1000);
                Element = driver.FindElement(By.Name("myName"));
                Element.SendKeys("Test Text");
                Thread.Sleep(2000);
                Console.WriteLine("Value : " + Element.GetAttribute("value"));
                Thread.Sleep(2000);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        [Test]
        public void CheckBox()
        {
            string[] Option = { "1", "3" };
            int i;
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/special-elements/check-button-test-3/");
                for (i = 0; i < Option.Length; i++)
                {
                    Element = driver.FindElement(By.CssSelector("#post-33 > div > p:nth-child(8) > input[type=\"checkbox\"]:nth-child(" + Option[i] + ")"));
                    Element.Click();
                    if (Element.GetAttribute("checked") == "true")
                    {
                        Console.WriteLine("Checkbox " + (i + 1) + " is Checked");
                        Console.WriteLine("Name : " + Element.GetAttribute("name"));
                        Console.WriteLine("Value : " + Element.GetAttribute("value"));
                    }
                    else
                    {
                        Console.WriteLine("Checkbox " + (i + 1) + " is not Checked");
                        Console.WriteLine("Name : " + Element.GetAttribute("name"));
                        Console.WriteLine("Value : " + Element.GetAttribute("value"));
                    }
                    Thread.Sleep(3000);
                }
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        [Test]
        public void RadioButton()
        {
            String[] Option = { "1", "3" , "5"};
            int i;
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/special-elements/radio-button-test/");
                for (i = 0; i < Option.Length; i++)
                {
                    Element = driver.FindElement(By.CssSelector(" #post-10 > div > form > p:nth-child(6) > input[type=\"radio\"]:nth-child(" +Option[i]+ ")"));
                    Element.Click();
                    if (Element.GetAttribute("checked") == "true")
                    {
                        Console.WriteLine("RadioButton " + (i + 1) + " Checked");
                        Console.WriteLine("Name :" + Element.GetAttribute("name"));
                        Console.WriteLine("Value " + Element.GetAttribute("value"));
                    }
                    else
                    {
                        Console.WriteLine("RadioButton " + (i + 1) + " not Checked");
                        Console.WriteLine("Name :" + Element.GetAttribute("name"));
                        Console.WriteLine("Value " + Element.GetAttribute("value"));
                    }
                }
                Thread.Sleep(3000);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        [Test]
        public void DropDown()
        {            
            int i;
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/special-elements/drop-down-menu-test/");
                Element = driver.FindElement(By.Name("DropDownTest"));
                Console.WriteLine("DropDown " + Element.GetAttribute("value") + " is selected");
                Console.WriteLine("DropDown Values are");
                for (i = 1; i <= 4; i++)
                {
                    Element = driver.FindElement(By.CssSelector("#post-6 > div > p:nth-child(6) > select > option:nth-child(" +i+ ")"));
                    Element.Click();                  
                    if (Element.GetAttribute("checked") == "true")
                    {
                        Console.WriteLine("DropDown " + i + " is Selected");
                        Console.WriteLine("Value : " + Element.GetAttribute("value"));
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        Console.WriteLine("Dropdown " + i + " is not Selected");
                    }
                }
                Thread.Sleep(3000);
            }
            catch (AssertionException e)
            {
                verificationErrors.Append(e.Message);
            }
        }

        [Test]
        public void AlertBox()
        {
            try
            {
                driver.Navigate().GoToUrl("http://testing.todvachev.com/special-elements/alert-box/");
                Alert = driver.SwitchTo().Alert();
                Thread.Sleep(3000);
                Console.WriteLine(Alert.Text);
                Alert.Accept();
                Element = driver.FindElement(By.CssSelector("#post-119 > div > figure > img"));
                if (Element.Displayed)
                {
                    Console.WriteLine("Alertbox Displayed");
                }                           
            }
            catch (AssertionException e)
            {
                Console.WriteLine("Alertbox not displayed");
                verificationErrors.Append(e.Message);
            }
        }

        [TearDown]
        public void CleanUp()
        {
            try
            {
                driver.Manage().Cookies.DeleteAllCookies();
                driver.Close();
                driver.Quit();
            }
            catch (Exception)
            {
                // Ignore errors if unable to close the browser
            }
            Assert.AreEqual("", verificationErrors.ToString());
        }
       
        private bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        private bool IsAlertPresent()
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException)
            {
                return false;
            }
        }

        private string CloseAlertAndGetItsText()
        {
            try
            {
                IAlert alert = driver.SwitchTo().Alert();
                string alertText = alert.Text;
                if (acceptNextAlert)
                {
                    alert.Accept();
                }
                else
                {
                    alert.Dismiss();
                }
                return alertText;
            }
            finally
            {
                acceptNextAlert = true;
            }
        }
    }
}