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
using OpenQA.Selenium;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;
using System.Drawing.Imaging;
using OpenQA.Selenium.Support.UI;

namespace Library
{
    /// <summary>
    /// Summary description for GeneralLibrary
    /// </summary>
    public class GeneralLibrary
    {
        public GeneralLibrary()
        {
        }

        public IWebDriver driver;

        #region GeneralFunctions
        //Function to check element present
        public bool isElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        //To wait until the element is present
        public void WaitunitleElementPresent(By by)
        {
            try
            {
                WebDriverWait wiat = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                wiat.Until((links) => { return links.FindElement(by); });
            }
            catch (Exception)
            {
            }
        }

        //To wait until the element is visible
        public void WaitunitleElementVisible(By by)
        {
            for (int i = 0; ; i++)
            {
                if (i >= 60) Assert.Fail("timeout");
                try
                {
                    if (driver.FindElement(by).Displayed) break;
                }
                catch (Exception)
                {
                }
            }
        }

        //Function to Read dynamic data into Excel sheet.
        public string Readdata(string fpath, int sheetno, int row, int col)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fpath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetno);
            range = xlWorkSheet.UsedRange;
            try
            {
                string dat = (string)(range.Cells[row, col] as Excel.Range).Value2;
                killprocess("EXCEL");
                return dat;
            }
            catch (Exception)
            {
                double dat = (double)(range.Cells[row, col] as Excel.Range).Value2;
                killprocess("EXCEL");
                return dat.ToString();
            }
        }

        //Kill Process
        public void killprocess(string pname)
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {

                if (clsProcess.ProcessName.StartsWith(pname))
                {
                    clsProcess.Kill();
                }
            }
        }

        //To get date & time stamp
        public string Time_Stamp()
        {
            string[] clCourse = new string[12];
            string FullTimeStamp;
            clCourse[0] = DateTime.Now.Day.ToString();
            clCourse[1] = "\\";
            clCourse[2] = DateTime.Now.Month.ToString();
            clCourse[3] = "\\";
            clCourse[4] = DateTime.Now.Year.ToString();
            clCourse[5] = "  ";
            clCourse[6] = DateTime.Now.Hour.ToString();
            clCourse[7] = ":";
            clCourse[8] = DateTime.Now.Minute.ToString();
            clCourse[9] = ":";
            clCourse[10] = DateTime.Now.Second.ToString();
            FullTimeStamp = String.Concat(clCourse);
            return FullTimeStamp;
        }

        //To get the unique name based on the time/date/seconds
        public string Generate_Unique_variable(string Unique)
        {
            string[] clCourse = new string[10];
            string fun_Unique;
            clCourse[0] = Unique;
            clCourse[1] = DateTime.Now.Year.ToString();
            clCourse[2] = DateTime.Now.Month.ToString();
            clCourse[3] = DateTime.Now.Day.ToString();
            clCourse[4] = DateTime.Now.Hour.ToString();
            clCourse[5] = DateTime.Now.Minute.ToString();
            clCourse[6] = DateTime.Now.Second.ToString();
            fun_Unique = String.Concat(clCourse);
            return fun_Unique;
        }

        //To create logs(html file) in drive
        public void Logs(string Logtext, string Status)
        {
            //To get OS drive
            //string drive = System.Environment.SystemDirectory.ToString();
            //string[] OS_Drive = drive.Split(':');
            string OS_Drive = "d";

            try
            {
                //To check result logs in OS drive
                string d = OS_Drive + ":\\Selenium_reports";
                if (!Directory.Exists(d))
                {
                    Directory.CreateDirectory(d);
                }
                bool fileExists1 = System.IO.File.Exists(OS_Drive + ":\\Selenium_reports\\Error_History.html");
                bool fileExists2 = System.IO.File.Exists(OS_Drive + ":\\Selenium_reports\\Passed_Results.html");
                bool fileExists3 = System.IO.File.Exists(OS_Drive + ":\\Selenium_reports\\Failed_Results.html");

                if (fileExists1 == false)
                {
                    System.IO.FileStream file1;
                    file1 = System.IO.File.Create(OS_Drive + ":\\Selenium_reports\\Error_History.html");
                    file1.Close();
                    System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Error_History.html", "<table border=\"1\"><tr><td width=33% align=\"center\">TEST SCENARIO</td><td width=33% align=\"center\">TIME</td><td width=33% align=\"center\">STATUS</td></tr>");
                    //file1.Close();
                }

                if (fileExists2 == false)
                {
                    System.IO.FileStream file1;
                    file1 = System.IO.File.Create(OS_Drive + ":\\Selenium_reports\\Passed_Results.html");
                    file1.Close();
                    System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Passed_Results.html", "<table border=\"1\"><tr><td width=33% align=\"center\">TEST SCENARIO</td><td width=33% align=\"center\">TIME</td><td width=33% align=\"center\">STATUS</td></tr>");

                }

                if (fileExists3 == false)
                {
                    System.IO.FileStream file1;
                    file1 = System.IO.File.Create(OS_Drive + ":\\Selenium_reports\\Failed_Results.html");
                    file1.Close();
                    System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Failed_Results.html", "<table border=\"1\"><tr><td width=33% align=\"center\">TEST SCENARIO</td><td width=33% align=\"center\">TIME</td><td width=33% align=\"center\">STATUS</td></tr>");

                }

                //To get last access date of result file
                string creation_DT = System.IO.File.GetLastAccessTime(OS_Drive + ":\\Selenium_reports\\Passed_Results.html").ToString();
                string current_DT = DateTime.Now.Date.ToString();
                string[] arr_DT = creation_DT.Split(' ');
                string[] Now_DT = current_DT.Split(' ');

                string time;
                string newtext;
                string newtext1;
                time = Time_Stamp();
                string mainText = System.IO.File.ReadAllText(OS_Drive + ":\\Selenium_reports\\Error_History.html");
                string maintext1 = ""; string maintext2 = "";

                //if date of access matches to current date then read all text inside the file. Else delete old logs and create new Passed and Failed logs
                if (arr_DT[0] == Now_DT[0])
                {
                    maintext1 = System.IO.File.ReadAllText(OS_Drive + ":\\Selenium_reports\\Passed_Results.html");
                    maintext2 = System.IO.File.ReadAllText(OS_Drive + ":\\Selenium_reports\\Failed_Results.html");
                }
                else
                {
                    System.IO.File.Delete(OS_Drive + ":\\Selenium_reports\\Passed_Results.html");
                    System.IO.File.Delete(OS_Drive + ":\\Selenium_reports\\Failed_Results.html");
                    System.IO.FileStream file1;
                    file1 = System.IO.File.Create(OS_Drive + ":\\Selenium_reports\\Passed_Results.html");
                    file1.Close();
                    System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Passed_Results.html", "<table border=\"1\"><tr><td width=33% align=\"center\">RESULT</td><td width=33% align=\"center\">TIME</td><td width=33% align=\"center\">Status</td></tr>");
                    System.IO.FileStream file2;
                    file2 = System.IO.File.Create(OS_Drive + ":\\Selenium_reports\\Failed_Results.html");
                    file2.Close();
                    System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Failed_Results.html", "<table border=\"1\"><tr><td width=33% align=\"center\">RESULT</td><td width=33% align=\"center\">TIME</td><td width=33% align=\"center\">Status</td></tr>");

                }

                if (mainText == "")
                {
                    //<font color=\"red\"> </font>
                    newtext = "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"blue\">" + Status + "</td></tr>";
                }
                else
                {
                    newtext = mainText + "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"blue\">" + Status + "</td></tr>";
                }
                System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Error_History.html", newtext);

                switch (Status)
                {
                    case "PASSED":
                        if (maintext1 == "")
                        {
                            newtext1 = "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"green\">" + Status + "</td></tr>";
                        }
                        else
                        {
                            //newtext1 = maintext1 + "\r\n\r\n" + Logtext + "      " + time + "    " + Status;
                            newtext1 = maintext1 + "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"green\">" + Status + "</td></tr>";
                        }
                        System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Passed_Results.html", newtext1);
                        break;

                    case "FAILED":
                        if (maintext2 == "")
                        {
                            newtext1 = "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"red\">" + Status + "</td></tr>";
                        }
                        else
                        {
                            newtext1 = maintext2 + "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"red\">" + Status + "</td></tr>";
                        }
                        System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Failed_Results.html", newtext1);
                        //Selenium.SelectWindow("");
                        string d1 = OS_Drive + ":\\Selenium_Screenshots";
                        if (!Directory.Exists(d1))
                        {
                            Directory.CreateDirectory(d1);
                        }
                        Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
                        ss.SaveAsFile(OS_Drive + ":\\Selenium_Screenshots\\" + Generate_Unique_variable("Error") + ".png", ImageFormat.Png);
                        //selenium.CaptureScreenshot(OS_Drive + ":\\Selenium_Screenshots\\" + Generate_Unique_variable("Error") + ".png");
                        break;

                    default:
                        if (maintext2 == "")
                        {
                            newtext1 = "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"red\">" + Status + "</td></tr>";
                        }
                        else
                        {
                            newtext1 = maintext2 + "<tr><td width=33%>" + Logtext + "</font></td><td width=33% align=\"center\"> " + time + "</td><td width=33% align=\"center\" bgcolor=\"red\">" + Status + "</td></tr>";
                        }
                        System.IO.File.WriteAllText(OS_Drive + ":\\Failed_Results.html", newtext1 + "Status not defined properly in the CODE");
                        break;
                }
            }
            catch (Exception e)
            {
                System.IO.File.WriteAllText(OS_Drive + ":\\Selenium_reports\\Failed_Results.html", e.Message + "FAILED");
            }

        }
        #endregion


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
