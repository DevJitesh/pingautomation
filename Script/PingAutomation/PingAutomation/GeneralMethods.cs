using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;

namespace PingAutomation
{
    class GeneralMethods
    {
        string Trackername = ConfigurationManager.AppSettings["Trackername"].ToString().ToUpper();
        string folderpth = ConfigurationManager.AppSettings["Folderpath"].ToString();

        Hashtable myHashtable;
        #region To read excel data
        public DataSet ImportexcelData(string FilePath,string ClientName)
        {

            OleDbConnection dbConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=\"Excel 12.0 xml;IMEX=1\";");
            dbConnection.Open();

            DataSet ds = new DataSet("Data");
            String sql = "Select * from [Sheet1$] where Agent_Name LIKE'" + ClientName+'%' + "'";
            OleDbCommand dbCommand = new System.Data.OleDb.OleDbCommand(sql, dbConnection);
            OleDbDataAdapter dbAdapter = new OleDbDataAdapter(dbCommand);

            DataTable dt = new DataTable("Sheet1");
            dbAdapter.Fill(dt);
            ds.Tables.Add(dt);

            dbConnection.Close();
            dbConnection.Dispose();

            //  CheckExcellProcesses();
            //  KillExcel();

            return ds;
        }
        #endregion

        #region To get list from tracker
        public DataSet isPresentontracker()
        {
            string connString=string.Empty;
            // int resquery = 0;
            if (Trackername.Contains("NEW"))
            {
                //new tracker 
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/Support/;LIST={DEFD5BEC-D6A2-400E-9F98-8F6E653A2CE9};VIEW={CC4F76AA-FA1B-44D7-9611-665077F0E451};"; //tracker new
            }
            else if (Trackername.Contains("OLD"))
            {
                //old tracker
                 connString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/;LIST={5D0E59F0-E498-4EF0-BBC1-C2FCF7C9150F};VIEW={87B39BBA-9465-4B2E-87D1-FE88E735EEA6};";   //tracker old
            }
            //string connString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/TIFFA-Kale/;LIST={27F0E118-ABC1-4D38-8FFD-2CB0C9F492D8};VIEW={E22AF388-5263-42B8-BB85-DE8B79BF6F61};";

            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();
            DataSet dsPingMe = new DataSet();

            // OdbcCommand command = new OdbcCommand("select * from Branch", conn);
            String h = "Hawb Count";                                                                                                
           OleDbCommand commandPingMe = new OleDbCommand("select * from LIST where Shipment_Status = '2 - Ack' AND `Hawb Count`=0", conn);
            //  OleDbCommand commandPingMe = new OleDbCommand("select Count(*)  from LIST where Pfx = '" + AirlinePfx + "' AND AWB_No = '" + MawbNo + "'", conn);
           // OleDbCommand commandPingMe = new OleDbCommand("select * from LIST", conn);
            OleDbDataAdapter dataAdapterPingMe = new OleDbDataAdapter(commandPingMe);
            DataTable dtPingMe = new DataTable();
            dataAdapterPingMe.Fill(dtPingMe);
            dsPingMe.Tables.Add(dtPingMe);
            //return Convert.ToInt32(commandPingMe.ExecuteScalar());
            //resquery = dsPingMe.Tables["Table1"].Rows.Count; 
            //return resquery;

            return dsPingMe;
        }
        #endregion

        public DataSet ImportexcelData1(string FilePath)
        {
           
                OleDbConnection dbConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=\"Excel 12.0 xml;IMEX=1\";");
                dbConnection.Open();

                DataSet ds = new DataSet("Data");
                String sql = "Select * from [Sheet1$]";
                OleDbCommand dbCommand = new System.Data.OleDb.OleDbCommand(sql, dbConnection);
                OleDbDataAdapter dbAdapter = new OleDbDataAdapter(dbCommand);

                DataTable dt = new DataTable("Sheet1");
                dbAdapter.Fill(dt);
                ds.Tables.Add(dt);

                dbConnection.Close();
                dbConnection.Dispose();

                //  CheckExcellProcesses();
                //  KillExcel();

                return ds;
        }
        
        public void CheckExcellProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process ExcelProcess in AllProcesses)
            {
                myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        public void KillExcel()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");

            // check to kill the right process
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (myHashtable.ContainsKey(ExcelProcess.Id) == true)
                    ExcelProcess.Kill();
            }

            AllProcesses = null;
        }

        #region Find Element xpath

        public bool TryFindElement(IWebDriver driver, string element1)
        {
            try
            {
                IWebElement element = driver.FindElement(By.XPath("" + element1 + ""));

                if (true)
                {
                    bool visible = IsElementVisible(element);

                    if (visible == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                //return true;
            }
            catch (NoSuchElementException ex)
            {
                return false;
                throw ex;
            }

        }

        public bool IsElementVisible(IWebElement element)
        {
            return element.Displayed && element.Enabled;
        }


        #endregion

        public void UpdatePingStatus(string AirlinePrefix, string MAWBNumber, StreamWriter strWriter)
        {
            try
            {

                System.Data.OleDb.OleDbConnection MyConnection=null;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                if (Trackername.Contains("NEW"))
                {
                    //new tracker
                     MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/Support/;LIST={DEFD5BEC-D6A2-400E-9F98-8F6E653A2CE9};VIEW={CC4F76AA-FA1B-44D7-9611-665077F0E451};");  //tracker new
                }
               else if (Trackername.Contains("OLD"))
                {
                    //old tracker
                    MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/;LIST={5D0E59F0-E498-4EF0-BBC1-C2FCF7C9150F};VIEW={87B39BBA-9465-4B2E-87D1-FE88E735EEA6};");  //tracker old
                }
                MyConnection.Open();

                myCommand.Connection = MyConnection;

                string DateStamp = System.DateTime.Now.ToString("dd/MM/yyyy H:mm");
                string sql = "Update [Table1] set Shipment_Status='4 - Updated in PING' where Pfx='" + AirlinePrefix + "' AND AWB_No='" + MAWBNumber + "' AND Shipment_Status='2 - Ack'";
                //, PDF_Sent_At = '" +DateStamp + "'
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                sql = "";
                MyConnection.Close();
                MyConnection.Dispose();
                Console.WriteLine("Status updated successfully for AWB " + MAWBNumber + ".\n");

                Console.WriteLine("Status updated successfully for AWB " + MAWBNumber + ".\n");

                strWriter.WriteLine(AirlinePrefix+"-"+MAWBNumber + " status Updated successfully as '4 - Updated in PING'");
                strWriter.WriteLine(" ");
                strWriter.WriteLine("-------------------------------------------------------");
                strWriter.WriteLine(" ");

                if (MyConnection.State == ConnectionState.Open)
                {
                    MyConnection.Close();
                    MyConnection.Dispose();
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                strWriter.WriteLine(MAWBNumber + " could not be Updated   " + ex.Message);
                strWriter.WriteLine(" ");
                strWriter.WriteLine("-------------------------------------------------------");
                strWriter.WriteLine(" ");

            }
        }

        public void changeoperator(string AirlinePrefix, string MAWBNumber,string operatorname, StreamWriter strWriter)
        {
            try
            {
                System.Data.OleDb.OleDbConnection MyConnection=null;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                if (Trackername.Contains("NEW"))
                {
                    //new tracker
                    MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/Support/;LIST={DEFD5BEC-D6A2-400E-9F98-8F6E653A2CE9};VIEW={CC4F76AA-FA1B-44D7-9611-665077F0E451};");  //tracker new
                    MyConnection.Open();
                }
               else if (Trackername.Contains("OLD"))
                {
                    //old tracker
                    MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/;LIST={5D0E59F0-E498-4EF0-BBC1-C2FCF7C9150F};VIEW={87B39BBA-9465-4B2E-87D1-FE88E735EEA6};");   //tracker old
                    MyConnection.Open();
                }
                    

                myCommand.Connection = MyConnection;

                string DateStamp = System.DateTime.Now.ToString("dd/MM/yyyy H:mm");
                string sql = "Update [Table1] set Operator='"+ operatorname + "' where Pfx='" + AirlinePrefix + "' AND AWB_No='" + MAWBNumber + "' AND Shipment_Status='2 - Ack'";
                //, PDF_Sent_At = '" +DateStamp + "'
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                sql = "";
                MyConnection.Close();
                MyConnection.Dispose();
                Console.WriteLine("Status updated successfully for AWB " + MAWBNumber + ".\n");

                Console.WriteLine("Status updated successfully for AWB " + MAWBNumber + ".\n");

                strWriter.WriteLine(AirlinePrefix + "-" + MAWBNumber + " status Updated successfully as '4 - Updated in PING'");
                strWriter.WriteLine(" ");
                strWriter.WriteLine("-------------------------------------------------------");
                strWriter.WriteLine(" ");

                if (MyConnection.State == ConnectionState.Open)
                {
                    MyConnection.Close();
                    MyConnection.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                strWriter.WriteLine(MAWBNumber + " could not be Updated   " + ex.Message);
                strWriter.WriteLine(" ");
                strWriter.WriteLine("-------------------------------------------------------");
                strWriter.WriteLine(" ");

            }
        }

        public void Assignedto(string AirlinePrefix, string MAWBNumber, string operatorname, StreamWriter strWriter)
        {
            try
            {
                System.Data.OleDb.OleDbConnection MyConnection=null;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                if (Trackername.Contains("NEW"))
                {
                    //new tracker                                                                                                                       
                     MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/Support/;LIST={DEFD5BEC-D6A2-400E-9F98-8F6E653A2CE9};VIEW={CC4F76AA-FA1B-44D7-9611-665077F0E451};");   //tracker new
                    MyConnection.Open();
                }
               else if (Trackername.Contains("OLD"))
                {
                    //old tracker
                    MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes; DATABASE=http://klsstg03/pingSD/;LIST={5D0E59F0-E498-4EF0-BBC1-C2FCF7C9150F};VIEW={87B39BBA-9465-4B2E-87D1-FE88E735EEA6};");   //tracker old
                    MyConnection.Open();
                }
                myCommand.Connection = MyConnection;

                string DateStamp = System.DateTime.Now.ToString("dd/MM/yyyy H:mm");
                string sql = "Update [Table1] set Assigned To='" + operatorname + "' where Pfx='" + AirlinePrefix + "' AND AWB_No='" + MAWBNumber + "' AND Shipment_Status='2 - Ack'";
                //, PDF_Sent_At = '" +DateStamp + "'
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                sql = "";
                MyConnection.Close();
                MyConnection.Dispose();
                Console.WriteLine("Status updated successfully for AWB " + MAWBNumber + ".\n");

                Console.WriteLine("Status updated successfully for AWB " + MAWBNumber + ".\n");

                strWriter.WriteLine(AirlinePrefix + "-" + MAWBNumber + " Assigned To Updated successfully as 'Jitesh Dev'");
                strWriter.WriteLine(" ");
                strWriter.WriteLine("-------------------------------------------------------");
                strWriter.WriteLine(" ");

                if (MyConnection.State == ConnectionState.Open)
                {
                    MyConnection.Close();
                    MyConnection.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                strWriter.WriteLine(MAWBNumber + " could not be Updated   " + ex.Message);
                strWriter.WriteLine(" ");
                strWriter.WriteLine("-------------------------------------------------------");
                strWriter.WriteLine(" ");

            }
        }

        #region alert

        public Boolean isAlertPresent(IWebDriver driver)
        {
            //bool result = false;
            // Boolean result1 = false;
            try
            {
                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                //IAlert alert1 = wait.Until(ExpectedConditions.AlertIsPresent());

                IAlert alert = driver.SwitchTo().Alert();
                Console.WriteLine("Alert present");
                //bool result = true;
                return true;
            }

            catch (NoAlertPresentException ex)
            {
                Console.WriteLine("Alert not present");
                //bool result = false;
                return false;
                throw ex;
            }
            // return result1;
        }

        #endregion

        public void selectOptionWithText(String textToSelect,IWebDriver driver)
        {
           
            try
            {
                IWebElement autoOptions = driver.FindElement(By.Id("ui-active-menuitem"));
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.Id("ui-active-menuitem"))));
                Thread.Sleep(300);
            

                IList<IWebElement> optionsToSelect = autoOptions.FindElements(By.TagName("a"));
                foreach (IWebElement option in optionsToSelect)
                {
                    if (option.Text.Equals(textToSelect))
                    {
                        Console.WriteLine("Trying to select: " + textToSelect);
                        option.Click();
                        break;
                    }
                }

            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.GetBaseException());
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        // creating folder datewise
        #region create folder datewise
        
            public string AutoFolderCreate(string MainFolderName)
            {
                string SavedFolder;
                SavedFolder = MainFolderName;
               
            //return excelpath1;

            bool folderexist = Directory.Exists(SavedFolder);
            string foldercheck = SavedFolder;

            if (folderexist != true)
            {
                Directory.CreateDirectory(folderpth + DateTime.Now.ToString("dd-MM-yyyy"));
            }
            else if (folderexist == true)
            {

            }
            string ExcelSavePath = AutoExcelCreate1(foldercheck);
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(foldercheck);
            int count = dir.GetFiles().Length;
            return ExcelSavePath;
        }

        #endregion
       

        public static string ExcelSavePath;
            //public string SavedFolder;
            public static string path1;
            public string AutoExcelCreate1(string SavedFolder)
            {
                string fnm =  DateTime.Now.ToString("dd-MM-yyyy");
                //Microsoft.Office.Interop.Excel.Application xlApp;
                //Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                //Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                //Excel.Range range;
                string ExcelSavePath = string.Empty;
                object misValue = System.Reflection.Missing.Value;

                int fCount = 0;
                fCount = Directory.GetFiles(SavedFolder, "*", SearchOption.AllDirectories).Length;

                //xlApp = new Excel.Application();
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Add();

                if (fCount != 0)
                {
                    for (int i = 0; i <= fCount; i++)
                    {
                        ExcelSavePath = System.IO.Path.Combine(SavedFolder, fnm + ".xls");
                        if (!File.Exists(ExcelSavePath))
                        {
                            
                        }
                    }
                }
                else
                {
                    ExcelSavePath = System.IO.Path.Combine(SavedFolder, fnm);
                    xlWorkSheet.Name = "Created_AWB";
                    xlWorkSheet.Cells[1, 1] = "Prefix";
                    xlWorkSheet.Cells[1, 2] = "AWB_No";
                   
                    xlWorkSheet.Cells[1, 1].Font.Bold = true;
                    //xlWorkSheet.Cells[1, 1].Interior.ColorIndex = 37;
                    xlWorkSheet.Cells[1, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[1, 2].Borders.LineStyle = true;
                    //xlWorkSheet.Cells[1, 2].Interior.ColorIndex = 37;
                  

                    xlWorkBook.SaveAs(ExcelSavePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);


                }
                return ExcelSavePath;
            }
        
            

            private static void releaseObject(object obj)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch (Exception)
                {
                    obj = null;
                }
                finally
                {
                    GC.Collect();
                }
            }
       



    }
}
