﻿using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Interactions;
using static PingAutomation.GeneralMethods;

namespace PingAutomation
{
    class create
    {
        string oper = null;
        public static void main()
        {
            IWebDriver driver=new ChromeDriver();
            StreamWriter strWriter;
                 FileStream files;
            String FileName = ConfigurationManager.AppSettings["LogfilePath"] + @"\PingAutomation" + DateTime.Now.ToString("dd-MMM-yyyy") + ".txt";
            files = new FileStream(FileName, FileMode.Append, FileAccess.Write, FileShare.None);
            strWriter = new StreamWriter(files);
            strWriter.WriteLine("---------------------------------------------------------------");
            strWriter.WriteLine("Log generated at " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm"));
            strWriter.WriteLine(" ");
            create c = new create();
            c.createawb(@"E:\JITESH\PingAutomation\Output_File\20190131_060021.xls", "MAGNUM CARGO PVT. LTD","020", "50752100", strWriter,driver);
        }

        public void createawb(string filename, string clientname, string Prefix, string awbno, StreamWriter strWriter,IWebDriver driver)
        {
            try
            {
                #region Webdriver initialization and login
                string[] Stringarr = null;
              
                bool result = false;
                string Trackernm = ConfigurationManager.AppSettings["Trackername"].ToString().ToUpper();
                if (Trackernm.Contains("NEW"))
                {
                    oper = ConfigurationManager.AppSettings["Operatornew"].ToString();
                }
                else if (Trackernm.Contains("OLD"))
                {
                    oper = ConfigurationManager.AppSettings["Operatorold"].ToString();
                }

                string url = ConfigurationManager.AppSettings["PingURL"].ToString();
                string username = ConfigurationManager.AppSettings["Username"].ToString();
                string password = ConfigurationManager.AppSettings["Password"].ToString();
                driver.Url = url;
                driver.FindElement(By.Id("txtUsrName")).SendKeys(username);
                driver.FindElement(By.Id("txtPswd")).SendKeys(password);
                driver.FindElement(By.Id("btnLogin")).Click();
                driver.Manage().Window.Maximize();
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='ctl00_lnkSignout']"))));
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//a[contains(text(),'AWB Service')]")).Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath("//a[contains(text(),'PING - Create MAWB')]"))));
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//a[contains(text(),'PING - Create MAWB')]")).Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='ctl00_hldPage_drpCopyFrom']"))));
                Thread.Sleep(1000);

                #endregion

                #region String read from excel
                GeneralMethods PTEGE = new GeneralMethods();
                DataSet d = new DataSet();

                string excelpath = filename;
                d = PTEGE.ImportexcelData1(filename);

                string
                  actual01 = null,
                  actual02 = null,
                  Shipper_Name = null,
                  Address_Line01 = null,
                  Address_Line02 = null,
                  Consignee_Name = null,
                  Address_Line01_c = null,
                  Address_Line02_c = null,
                  Agent = null,
                  Accounting_Information = null,
                  Origin_Port,
                  Destn_Port,
                  Via01,
                  Via02, Chargeable_Weight,
                  Charge_code = null,
                  Flight_Details_01 = null,
                  Flight_Date_01 = null,
                  Flight_Details_02 = null,
                  Flight_Date_02 = null,
                  Handling_Information = null,
                  No_pcs = null,
                  Gross_Wt = null,
                  Rate_Class = null,
                  Charges = null,
                  Commodity_No = null,
                  Nature = null,
                Dimension = null,
    Executed_On_Date = null,
    Total=null,
    At_Place = null;



                string Prefixdt = d.Tables["Sheet1"].Rows[0]["pre"].ToString();

                actual01 = d.Tables["Sheet1"].Rows[0]["actual01"].ToString();
                actual02 = d.Tables["Sheet1"].Rows[0]["actual02"].ToString();
                Shipper_Name = d.Tables["Sheet1"].Rows[0]["Shipper_Name"].ToString();
                Address_Line01 = d.Tables["Sheet1"].Rows[0]["Address_Line01"].ToString();
                Address_Line02 = d.Tables["Sheet1"].Rows[0]["Address_Line02"].ToString();
                Consignee_Name = d.Tables["Sheet1"].Rows[0]["Consignee_Name"].ToString();
                Address_Line01_c = d.Tables["Sheet1"].Rows[0]["Address_Line01_c"].ToString();
                Address_Line02_c = d.Tables["Sheet1"].Rows[0]["Address_Line02_c"].ToString();
                Agent = d.Tables["Sheet1"].Rows[0]["Agent"].ToString();
                Accounting_Information = d.Tables["Sheet1"].Rows[0]["Accounting_Information"].ToString();
                Origin_Port = d.Tables["Sheet1"].Rows[0]["Origin_Port"].ToString();
                Destn_Port = d.Tables["Sheet1"].Rows[0]["Destn_Port"].ToString();
                Via01 = d.Tables["Sheet1"].Rows[0]["Via01"].ToString();
                Via02 = d.Tables["Sheet1"].Rows[0]["Via02"].ToString();
                Charge_code = d.Tables["Sheet1"].Rows[0]["Charge_code"].ToString();
                Flight_Details_01 = d.Tables["Sheet1"].Rows[0]["Flight_Details_01"].ToString();
                Flight_Date_01 = d.Tables["Sheet1"].Rows[0]["Flight_Date_01"].ToString();
                if (d.Tables.Contains("Flight_Details_02"))
                {
                    Flight_Details_02 = d.Tables["Sheet1"].Rows[0]["Flight_Details_02"].ToString();
                    Flight_Date_02 = d.Tables["Sheet1"].Rows[0]["Flight_Date_02"].ToString();
                }
                Handling_Information = d.Tables["Sheet1"].Rows[0]["Handling_Information"].ToString();
                No_pcs = d.Tables["Sheet1"].Rows[0]["No_pcs"].ToString();
                Gross_Wt = d.Tables["Sheet1"].Rows[0]["Gross_Wt"].ToString();
                Rate_Class = d.Tables["Sheet1"].Rows[0]["Rate_Class"].ToString();
                Chargeable_Weight = d.Tables["Sheet1"].Rows[0]["Chargeable_Weight"].ToString();
                Charges = d.Tables["Sheet1"].Rows[0]["Charges"].ToString();
                Commodity_No = d.Tables["Sheet1"].Rows[0]["Commodity_No"].ToString();
                Nature = d.Tables["Sheet1"].Rows[0]["Nature"].ToString();

                Dimension = d.Tables["Sheet1"].Rows[0]["Dimension"].ToString();
                Executed_On_Date = d.Tables["Sheet1"].Rows[0]["Executed_On_Date"].ToString();
                At_Place = d.Tables["Sheet1"].Rows[0]["At_Place"].ToString();
                Total = d.Tables["Sheet1"].Rows[0]["Total"].ToString();
                #endregion

                #region code for create awb
             //   awbno = actual01 + actual02;

                IWebElement awbpre = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtAWBPrefix']"));
              awbpre.SendKeys(Prefix);
              // awbpre.SendKeys(Prefixdt);
                Thread.Sleep(500);
                awbpre.SendKeys(Keys.Control + "a");
                Thread.Sleep(500);
                awbpre.SendKeys(Keys.Tab);
                Thread.Sleep(500);
                IWebElement awbn = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtAWBNo']"));
                awbn.SendKeys(awbno);
                awbn.SendKeys(Keys.Tab);

                Thread.Sleep(2000);
                result = PTEGE.TryFindElement(driver, ".//*[@id='ctl00_hldPage_lblnotify']");
                if (result == true)
                {
                    string errmsg = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_lblnotify']")).Text;
                    if (errmsg.Contains("is already created please use different AWB No"))
                    {
                        PTEGE.changeoperator(Prefix, awbno, oper, strWriter);
                        PTEGE.UpdatePingStatus(Prefix, awbno, strWriter);                        
                    }
                }

                driver.FindElement(By.XPath(".//*[@id='btnshipperadd']")).Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='ctl00_hldPage_txtOrgName']"))));
                Thread.Sleep(300);
                Shipper_Name = Shipper_Name.Replace("[^a-zA-Z0-9]", " ");
                string Shipper_Name_01 = Regex.Replace(Shipper_Name, @"[^0-9a-zA-Z]+", " ");

                // string Shipper_Name_01 = Shipper_Name.Replace("( +.,-&:/)", " ").Trim();
                driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtOrgName']")).SendKeys(Shipper_Name_01);

                driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtCompanyName']")).SendKeys(Shipper_Name_01);

                string addln1splited = null;
                Address_Line01 = Address_Line01.Replace("[^a-zA-Z0-9]", " ");
                string Address_Line01s = Regex.Replace(Address_Line01, @"[^0-9a-zA-Z]+", " ");
                Thread.Sleep(400);
                int shpaddline1cnt = Address_Line01s.Length;
                if (shpaddline1cnt > 35)
                {
                    Thread.Sleep(400);
                    int lll = shpaddline1cnt - 34;
                    string addline1spl = Address_Line01s.Substring(0, 34);
                    addln1splited = Address_Line01s.Substring(34, lll);
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).SendKeys(addline1spl);
                }
                else
                {
                    Thread.Sleep(400);
                    //  string Address_Line01s = Address_Line01.Replace("( +.,-&:/)", " ").Trim();
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).SendKeys(Address_Line01s);
                }


                Address_Line02 = Address_Line02.Replace("[^a-zA-Z0-9]", " ");
                string Address_Line02s = Regex.Replace(Address_Line02, @"[^0-9a-zA-Z]+", " ");
                string addln2new = addln1splited + Address_Line02s;
                Thread.Sleep(400);
                int addln2cntcn = addln2new.Length;
                if (addln2cntcn > 35)
                {
                    Thread.Sleep(400);
                    string splitedaddline2 = addln2new.Substring(0, 35);
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine2']")).SendKeys(splitedaddline2);
                }
                else
                {
                    Thread.Sleep(400);
                    //string Address_Line02s = Address_Line02.Replace("( +.,-&:/)", " ").Trim();
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine2']")).SendKeys(addln2new);
                }


                IWebElement cntname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCountry_txtName']"));
                cntname.SendKeys("India");
                cntname.SendKeys(Keys.Control + "a");
                cntname.SendKeys(Keys.Tab);

                IWebElement genauttxt = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCity_txtName']"));
                genauttxt.SendKeys("MUMBAI");
                genauttxt.SendKeys(Keys.Tab);
              //  driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtOtherCity']")).SendKeys("MUM");
                driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtPinCode']")).SendKeys("xxxxxx");
                IWebElement uld1 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtPinCode']"));
                uld1.Click();
                uld1.SendKeys(Keys.Tab + Keys.Enter);


                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='imgbtnConsignee']"))));
                Thread.Sleep(300);
                driver.FindElement(By.XPath(".//*[@id='imgbtnConsignee']")).Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='ctl00_hldPage_txtSearchConsigneeName']"))));
                Thread.Sleep(300);
                Consignee_Name = Consignee_Name.Replace("[^a-zA-Z0-9]", " ");
                string Consignee_Names = Regex.Replace(Consignee_Name, @"[^0-9a-zA-Z]+", " ");
                // string Consignee_Names = Consignee_Name.Replace("( +.,-&:/)", " ").Trim();
                driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtSearchConsigneeName']")).SendKeys(Consignee_Names);
                driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtConName']")).SendKeys(Consignee_Names);

                Address_Line01_c = Address_Line01_c.Replace("[^a-zA-Z0-9]", " ");
                string Address_Line01_cs = Regex.Replace(Address_Line01_c, @"[^0-9a-zA-Z]+", " ");
                Thread.Sleep(400);
                int addlin1cnt = Address_Line01_cs.Length;
                string addlin1cons = null;
                string Address_Line02_cs = null;
                if (addlin1cnt > 35)
                {
                    Thread.Sleep(400);
                    int lll = addlin1cnt - 34;
                    string consaddln1 = Address_Line01_cs.Substring(0, 34);
                    addlin1cons = Address_Line01_cs.Substring(34, lll);
                    
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine1']")).SendKeys(consaddln1);

                }
                else
                {
                    Thread.Sleep(400);
                    // string Address_Line01_cs = Address_Line01_c.Replace("( +.,-&:/)", " ").Trim();
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine1']")).SendKeys(Address_Line01_cs);
                }

                Address_Line02_c = Address_Line02_c.Replace("[^a-zA-Z0-9]", " ");
                Address_Line02_cs = Regex.Replace(Address_Line02_c, @"[^0-9a-zA-Z]+", " ");
                string addline2join = addlin1cons + Address_Line02_cs;
                Thread.Sleep(400);
                int addln2cnt = addline2join.Length;
                if (addln2cnt > 35)
                {
                    Thread.Sleep(400);
                    string splitaddline2 = addline2join.Substring(0, 35);
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine2']")).SendKeys(splitaddline2);
                }
                else
                {
                    Thread.Sleep(400);
                    //  string Address_Line02_cs = Address_Line02_c.Replace("( +.,-&:/)", " ").Trim();
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine2']")).SendKeys(addline2join);
                }

                IWebElement cons = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCountry_txtName']"));
                cons.SendKeys("germany");
                cons.SendKeys(Keys.Control + "a");
                cons.SendKeys(Keys.ArrowDown);
                cons.SendKeys(Keys.Tab);

                IWebElement conscity = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCity_txtName']"));
                conscity.SendKeys("Frankfurt");
                conscity.SendKeys(Keys.Control + "a");
                conscity.SendKeys(Keys.ArrowDown);
                conscity.SendKeys(Keys.Tab);

                driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtPinCode']")).SendKeys("xxxxxx");
                IWebElement uld = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_uclAlsoNotifyAddressInfo_txtPinCode']"));
                uld.Click();
                uld.SendKeys(Keys.Tab + Keys.Enter);


                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='ctl00_hldPage_txtAgentSelect']"))));
                Thread.Sleep(300);
                            
                string agent1 = Agent.Substring(0, 17);

                DataSet ddd = new DataSet();
                string Issuing_Agent_master = ConfigurationManager.AppSettings["Issuing_Agent"].ToString();
                ddd = PTEGE.ImportexcelData(Issuing_Agent_master, agent1);
                int v = ddd.Tables[0].Rows.Count;
                for (int s = 0; s < v; s++)
                {
                    string cnm = ddd.Tables["Sheet1"].Rows[s]["Agent_Name"].ToString();
                    string branchname = ddd.Tables["Sheet1"].Rows[s]["Branch"].ToString();
                    string mastername = ddd.Tables["Sheet1"].Rows[s]["Issuing_Agent_Name"].ToString();
                    mastername = mastername.Substring(0, 10);
                    string agnt = Agent.ToUpper();

                    string bnm = branchname.ToUpper();
                    if (bnm.Contains("BANGALORE"))
                    {
                        branchname = "Bengaluru";
                    }
                    if (bnm.Contains("BANGLORE"))
                    {
                        branchname = "Bengaluru";
                    }

                    if (bnm.Contains("BOM"))
                    {
                        branchname = "Mumbai";
                    }
                   
                    if (agnt.Contains(bnm))
                   {
                        IWebElement txtAgentSelect = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtAgentSelect']"));
                        txtAgentSelect.SendKeys(mastername);
                        txtAgentSelect.SendKeys(Keys.Control + "a");
                        Thread.Sleep(3000);
                        Actions act = new Actions(driver);
                        IWebElement ct = driver.FindElement(By.XPath("//a[contains(text(),'"+ branchname+"')]"));
                        Thread.Sleep(3000);
                        act.MoveToElement(ct).Click();
                        ct.Click();                  
                        Thread.Sleep(3000);
                        s = v;                    
                  }
                }
                Accounting_Information = Accounting_Information.Replace("[^a-zA-Z0-9]", " ");
                string Accounting_Information_01 = Regex.Replace(Accounting_Information, @"[^0-9a-zA-Z]+", " ");
                // string Accounting_Information_01 = Accounting_Information.Replace("( +.,-&:/)", " ").Trim();
                driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_accountinginfo']")).SendKeys(Accounting_Information_01);

                //   declare string as null

                string Via1;
                Via1 = Via01;
                string Via1_01 = Via1.Replace("( +.,-&:/)", " ").Trim();
                if (string.IsNullOrEmpty(Via1_01) || string.IsNullOrEmpty(Via1))
                {
                    //  Via1_01 = "null";

                }

                string Via2;
                Via2 = Via02;
                string Via2_01 = Via2.Replace("( +.,-&:/)", " ").Trim();

                if (string.IsNullOrEmpty(Via2_01) || string.IsNullOrEmpty(Via2))
                {
                    // Via2_01 = "null";
                }

                string DestPort;
                DestPort = Destn_Port;
                string DestPort_01 = DestPort.Replace("( +.,-&:/)", " ").Trim();
                if (string.IsNullOrEmpty(DestPort_01) || string.IsNullOrEmpty(DestPort))
                {
                    // DestPort_01 = "null";
                }


                #region Routing details

                string Airport_City_Master = ConfigurationManager.AppSettings["Airport_City_Master"].ToString();               
                DataSet dtset_air = new DataSet();
                dtset_air = PTEGE.ImportexcelData1(Airport_City_Master);
                int ss = dtset_air.Tables[0].Rows.Count;
                
                string Current_Date_for_defaultflight = DateTime.Now.ToString("dd");
                driver.FindElement(By.XPath(".//*[@id='imgairport']")).Click();
                int orport = Origin_Port.Length;
                if (orport <= 3)
                {
                    IWebElement airportcode = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillOriginAirport_txtCode']"));
                    airportcode.SendKeys(Origin_Port);
                    airportcode.SendKeys(Keys.Tab);
                    result = PTEGE.isAlertPresent(driver);
                    if (result == true)
                    {
                        driver.SwitchTo().Alert().Accept();
                        for (int lr = 0; lr < ss; lr++)
                        {
                            string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                            string name= dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString();
                            if (Origin_Port.Contains(code))
                            {
                                IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillOriginAirport_txtName']"));
                                airportname.SendKeys(name);
                                airportname.SendKeys(Keys.Tab);
                                lr = ss;
                            }
                        }

                      //  IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillOriginAirport_txtName']"));
                     //   airportname.SendKeys("London");
                     //   airportname.SendKeys(Keys.Tab);
                    }
                }
                else
                {
                    string oport01 = Origin_Port.Substring(0, 4);
                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillOriginAirport_txtName']"));
                    airportname.SendKeys(oport01);
                    airportname.SendKeys(Keys.Tab);
                    result = PTEGE.isAlertPresent(driver);
                    if (result == true)
                    {
                        driver.SwitchTo().Alert().Accept();
                        IWebElement airportname1 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillOriginAirport_txtName']"));
                        airportname1.SendKeys("London");
                        airportname1.SendKeys(Keys.Tab);
                    }
                }

                if (!string.IsNullOrEmpty(Destn_Port))//check empty
                {
                    /* if (Destn_Port.Contains("LON"))
                     {
                             IWebElement destprt = driver.FindElement(By.Id("ctl00_hldPage_GenericAutoFillDestAirport_txtName"));
                             destprt.SendKeys("London");
                             destprt.SendKeys(Keys.Control + "a");
                             destprt.SendKeys(Keys.ArrowDown);
                             destprt.SendKeys(Keys.Tab);
                     }
                     else
                     {*/
                    int destport = Destn_Port.Length;
                    if (destport <= 3)
                    {
                        IWebElement destpo = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']"));
                        destpo.SendKeys(Destn_Port);
                        destpo.SendKeys(Keys.Tab);

                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            driver.SwitchTo().Alert().Accept();
                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString();
                                if (Destn_Port.Contains(code))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                        }
                    }
                    else
                    {
                        IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtName']"));
                        airportname.SendKeys(Destn_Port);
                        airportname.SendKeys(Keys.Tab);

                    }
                    //  }


                    /*
                    if (Via01.Contains("LON"))
                    {
                        IWebElement via1prt = driver.FindElement(By.Id("ctl00_hldPage_ViaRoute1_txtName"));
                        via1prt.SendKeys("London");
                        via1prt.SendKeys(Keys.Control + "a");
                        via1prt.SendKeys(Keys.ArrowDown);
                        via1prt.SendKeys(Keys.Tab);
                    }
                    else
                    {
                    */
                    int via1p = Via01.Length;
                    if (via1p <= 3)
                    {
                        IWebElement via1prt = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute1_txtCode']"));
                        via1prt.SendKeys(Via01);
                        via1prt.SendKeys(Keys.Tab);
                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            driver.SwitchTo().Alert().Accept();
                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString();
                                if (Via01.Contains(code))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute1_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                        }
                    }
                    else
                    {
                        IWebElement destprt1 = driver.FindElement(By.Id("ctl00_hldPage_ViaRoute1_txtName"));
                        destprt1.SendKeys(Via01);
                        destprt1.SendKeys(Keys.Control + "a");
                        destprt1.SendKeys(Keys.ArrowDown);
                        destprt1.SendKeys(Keys.Tab);
                    }

                    //}

                    /*
                    if (Via02.Contains("LON"))
                    {
                        IWebElement btn_via2 = driver.FindElement(By.Id("ctl00_hldPage_ViaRoute2_txtName"));
                        btn_via2.SendKeys("London");
                        btn_via2.SendKeys(Keys.Control + "a");
                        btn_via2.SendKeys(Keys.ArrowDown);
                        btn_via2.SendKeys(Keys.Tab);
                    }
                    else
                    {*/

                    int via2lng = Via02.Length;
                    if (via2lng <= 3)
                    {
                        IWebElement btn_via2 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute2_txtCode']"));
                        btn_via2.SendKeys(Via02);
                        btn_via2.SendKeys(Keys.Tab);
                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            driver.SwitchTo().Alert().Accept();
                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString();
                                if (Via02.Contains(code))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute2_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                        }
                    }
                    else
                    {
                        IWebElement destprt1 = driver.FindElement(By.Id("ctl00_hldPage_ViaRoute2_txtName"));
                        destprt1.SendKeys(Via02);
                        destprt1.SendKeys(Keys.Control + "a");
                        destprt1.SendKeys(Keys.ArrowDown);
                        destprt1.SendKeys(Keys.Tab);
                    }


                    //  }
                    result = PTEGE.TryFindElement(driver, ".//*[@id='ctl00_hldPage_ViaRoute2_txtName']");
                    if (result == true)
                    {
                        IWebElement via_2 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute2_txtName']"));
                        via_2.SendKeys(Keys.Tab + Keys.Enter);
                    }


                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtflightno1']")).SendKeys("123");
                    string tomorrow = DateTime.Now.AddDays(1).ToString("dd/MM/yyyy");
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtflightdate1']")).SendKeys(tomorrow);


                }
                else if (!string.IsNullOrEmpty(Via2_01))
                {
                    /*
                    if (Via2_01.Contains("LON"))
                    {
                        IWebElement btn_via2 = driver.FindElement(By.Id("ctl00_hldPage_GenericAutoFillDestAirport_txtName"));
                        btn_via2.SendKeys("London");
                        btn_via2.SendKeys(Keys.Control + "a");
                        btn_via2.SendKeys(Keys.ArrowDown);
                        btn_via2.SendKeys(Keys.Tab);
                    }
                    else
                    {*/
                    int via2p = Via2_01.Length;
                    if (via2p <= 3)
                    {
                        IWebElement btn_via21 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']"));
                        btn_via21.SendKeys(Via2_01);
                        btn_via21.SendKeys(Keys.Tab);
                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            driver.SwitchTo().Alert().Accept();
                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString();
                                if (Via2_01.Contains(code))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                        }
                    }
                    else
                    {
                        IWebElement destprt = driver.FindElement(By.Id("ctl00_hldPage_GenericAutoFillDestAirport_txtName"));
                        destprt.SendKeys(Via2_01);
                        destprt.SendKeys(Keys.Control + "a");
                        destprt.SendKeys(Keys.ArrowDown);
                        destprt.SendKeys(Keys.Tab);
                    }


                    //  }
                    /*
                    if (Via1_01.Contains("LON"))
                    {
                        IWebElement btn_via2 = driver.FindElement(By.Id("ctl00_hldPage_ViaRoute1_txtName"));
                        btn_via2.SendKeys("London");
                        btn_via2.SendKeys(Keys.Control + "a");
                        btn_via2.SendKeys(Keys.ArrowDown);
                        btn_via2.SendKeys(Keys.Tab);
                    }
                    else
                    {*/
                    int viart1p = Via1_01.Length;
                    if (viart1p <= 3)
                    {
                        IWebElement viaroute1 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute1_txtCode']"));
                        viaroute1.SendKeys(Via1_01);
                        viaroute1.SendKeys(Keys.Tab);
                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            driver.SwitchTo().Alert().Accept();
                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString();
                                if (Via1_01.Contains(code))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute1_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                        }
                    }
                    else
                    {
                        IWebElement destprt = driver.FindElement(By.Id("ctl00_hldPage_ViaRoute1_txtName"));
                        destprt.SendKeys(Via1_01);
                        destprt.SendKeys(Keys.Control + "a");
                        destprt.SendKeys(Keys.ArrowDown);
                        destprt.SendKeys(Keys.Tab);
                    }
                    result = PTEGE.TryFindElement(driver, ".//*[@id='ctl00_hldPage_ViaRoute2_txtName']");
                    if (result == true)
                    {
                        IWebElement via_2 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute2_txtName']"));
                        via_2.SendKeys(Keys.Tab + Keys.Enter);
                    }
                    //change keystab
                    //driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[9]/div[11]/button[1]")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtflightno1']")).SendKeys("123");
                    string tomorrow = DateTime.Now.AddDays(1).ToString("dd/MM/yyyy");
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtflightdate1']")).SendKeys(tomorrow);

                }
                else
                {
                    /*
                    if (Via1.Contains("LON"))
                    {
                        IWebElement btn_via2 = driver.FindElement(By.Id("ctl00_hldPage_GenericAutoFillDestAirport_txtName"));
                        btn_via2.SendKeys("London");
                        btn_via2.SendKeys(Keys.Control + "a");
                        btn_via2.SendKeys(Keys.ArrowDown);
                        btn_via2.SendKeys(Keys.Tab);
                    }
                    else
                    {*/
                    int via11ln = Via1.Length;
                    if (via11ln <= 3)
                    {
                        IWebElement via1 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']"));
                        via1.SendKeys(Via1);
                        via1.SendKeys(Keys.Tab);
                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            driver.SwitchTo().Alert().Accept();
                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString();
                                if (Via1.Contains(code))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                        }                      
                    }
                    else
                    {
                        IWebElement destprt = driver.FindElement(By.Id("ctl00_hldPage_GenericAutoFillDestAirport_txtName"));
                        destprt.SendKeys(Via1);
                        destprt.SendKeys(Keys.Tab);
                    }
                    //}
                    result = PTEGE.TryFindElement(driver, ".//*[@id='ctl00_hldPage_ViaRoute2_txtName']");
                    if (result == true)
                    {
                        IWebElement via_2 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ViaRoute2_txtName']"));
                        via_2.SendKeys(Keys.Tab + Keys.Enter);
                    }

                    result = PTEGE.TryFindElement(driver, ".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtName']");
                    if (result == true)
                    {
                        IWebElement via_2 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtName']"));
                        via_2.SendKeys(Keys.Tab + Keys.Enter);
                    }
                    // driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[9]/div[11]/button[1]")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtflightno1']")).SendKeys("123");
                    string tomorrow = DateTime.Now.AddDays(1).ToString("dd/MM/yyyy");
                    driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtflightdate1']")).SendKeys(tomorrow);
                }

                #endregion

                #region chargecode
                if (!string.IsNullOrEmpty(Charge_code))
                {
                    if (Charge_code.Contains("PP") || Charge_code.Equals("P"))
                    {
                        SelectElement chgcode = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ddlChargeCode']")));
                        chgcode.SelectByText("PX");
                    }
                    else if (Charge_code.Contains("CX") || Charge_code.Equals("C"))
                    {
                        SelectElement chgcode = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ddlChargeCode']")));
                        chgcode.SelectByText("PX");
                    }
                    else
                    {
                        SelectElement chgcode = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ddlChargeCode']")));
                        chgcode.SelectByText(Charge_code);
                    }
                }
                else
                {
                    SelectElement chgcode = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_ddlChargeCode']")));
                    chgcode.SelectByText("PX");
                }
                #endregion
                Handling_Information = Handling_Information.Replace("[^a-zA-Z0-9]", " ");
                string HandlingInfo_01 = Regex.Replace(Handling_Information, @"[^0-9a-zA-Z]+", " ");
                //  string HandlingInfo_01 = Handling_Information.Replace("( +.,-&:/)", " ").Trim();
                IWebElement handling = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtssr']"));
                handling.Clear();
                handling.SendKeys(HandlingInfo_01);

                #region dimension
                //click on image dimension
                driver.FindElement(By.XPath(".//*[@id='addDimensions_1']")).Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='txtNoPcs_1']"))));
                Thread.Sleep(300);
                string l = "10", w = "10", h = "10";
                /*
                if (!string.IsNullOrEmpty(Dimension))
                {
                    string dim = Dimension.Substring(11);
                    dim = dim.Replace(" ", "");

                    string[] splchr = dim.Split(',');
                    int lng = splchr.Length;
                    int s = 1;
                    for (int q = 0; q < lng; q++)
                    {
                        s = q + 1;
                        string strarr = Convert.ToString(splchr[q]);
                        string[] dimn = null;
                        dimn = strarr.Split('=');
                        No_pcs = Convert.ToString(dimn[0]);
                        string dimens = Convert.ToString(dimn[1]);
                        string[] lwh = dimens.Split('x');
                        l = Convert.ToString(lwh[0]);
                        w = Convert.ToString(lwh[1]);
                        h = Convert.ToString(lwh[2]);
                        driver.FindElement(By.XPath(".//*[@id='txtNoPcs_" + s + "']")).SendKeys(No_pcs);

                        driver.FindElement(By.XPath(".//*[@id='txtLength_" + s + "']")).SendKeys(l);
                        driver.FindElement(By.XPath(".//*[@id='txtWidth_" + s + "']")).SendKeys(w);
                        driver.FindElement(By.XPath(".//*[@id='txtHeight_" + s + "']")).SendKeys(h);


                        if (s < lng)
                        {
                            int v = s + 1;
                            driver.FindElement(By.XPath(".//*[@id='addrow_" + s + "']")).Click();
                            new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath(".//*[@id='txtNoPcs_" + v + "']"))));
                            Thread.Sleep(300);
                        }
                    }
                    result = PTEGE.TryFindElement(driver, ".//*[@id='txtULDOwnerCode_2']");
                    if (result == true)
                    {
                        IWebElement uldo = driver.FindElement(By.XPath(".//*[@id='txtULDOwnerCode_" + s + "']"));
                        uldo.Click();
                        uldo.SendKeys(Keys.Tab + Keys.Tab + Keys.Tab + Keys.Enter);
                    }
                    else
                    {
                        IWebElement uldo = driver.FindElement(By.XPath(".//*[@id='txtULDOwnerCode_1']"));
                        uldo.Click();
                        uldo.SendKeys(Keys.Tab + Keys.Tab + Keys.Enter);
                    }
                }
                else
                {
*/
                    driver.FindElement(By.XPath(".//*[@id='txtNoPcs_1']")).SendKeys(No_pcs);

                    driver.FindElement(By.XPath(".//*[@id='txtLength_1']")).SendKeys(l);
                    driver.FindElement(By.XPath(".//*[@id='txtWidth_1']")).SendKeys(w);
                    driver.FindElement(By.XPath(".//*[@id='txtHeight_1']")).SendKeys(h);
                    IWebElement uldo = driver.FindElement(By.XPath(".//*[@id='txtULDOwnerCode_1']"));
                    uldo.Click();
                    uldo.SendKeys(Keys.Tab + Keys.Tab + Keys.Enter);

              //  }
                #endregion
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementExists((By.XPath(".//*[@id='txtCgGrWt_1']"))));
                Thread.Sleep(500);

                driver.FindElement(By.XPath(".//*[@id='txtCgGrWt_1']")).SendKeys(Gross_Wt);


                #region rateclass
                string RateClass = null;
                RateClass = Rate_Class;
                string RateClass_01 = null;
                RateClass_01 = Regex.Replace(RateClass, @"[^0-9a-zA-Z]+", " ");
                // RateClass_01 = RateClass.Replace("( +.,-&:/)", " ").Trim();
                if (string.IsNullOrEmpty(RateClass_01))
                {
                    RateClass_01 = "null";
                }

                if (RateClass_01 != "null")
                {
                    SelectElement rc1 = new SelectElement(driver.FindElement(By.XPath(".//*[@id='selRateClass_1']")));
                    rc1.SelectByText(RateClass_01);

                }
                else
                {
                    SelectElement rc1 = new SelectElement(driver.FindElement(By.XPath(".//*[@id='selRateClass_1']")));
                    rc1.SelectByText("Q");
                }
                #endregion

                #region commodity no

                string ComdityNo_01 = null;
                ComdityNo_01 = Regex.Replace(Commodity_No, @"[^0-9a-zA-Z]+", " ");
                // ComdityNo_01 =Commodity_No.Replace("( +.,-&:/)", " ").Trim();

                if (Rate_Class.Contains("C") || Rate_Class.Contains("S"))
                {
                    if (Commodity_No == "null" || Commodity_No == " ")
                    {
                        driver.FindElement(By.XPath(".//*[@id='txtCommNo_1']")).SendKeys("111");
                    }
                    else
                    {
                        driver.FindElement(By.XPath(".//*[@id='txtCommNo_1']")).SendKeys(ComdityNo_01);
                    }

                }
                #endregion

                //==chargeable wt
                IWebElement chgwt = driver.FindElement(By.Id("txtCgChargWt_1"));
                chgwt.Clear();
                chgwt.SendKeys(Chargeable_Weight);
                chgwt.SendKeys(Keys.Tab);



                #region charges

                string Charge_01 = null;
                Charge_01 = Regex.Replace(Charges, @"[^0-9a-zA-Z]+", " ");
                // Charge_01 = Charges.Replace("( +.,-&:/)", " ").Trim();
                string chg1 = Charge_01.ToUpper();
                if (chg1.Contains("MIN")|| Rate_Class.Equals("M"))
                {
                    driver.FindElement(By.XPath(".//*[@id='txtCgRate_1']")).SendKeys(Total);
                    Thread.Sleep(1000);                    
                }
                else
                {
                    if (Charge_01.Contains("null") || string.IsNullOrEmpty(Charge_01))
                    {
                        driver.FindElement(By.XPath(".//*[@id='txtCgRate_1']")).SendKeys("1");
                    }
                    else
                    {
                        driver.FindElement(By.XPath(".//*[@id='txtCgRate_1']")).SendKeys(Charge_01);
                        Thread.Sleep(1000);
                    }
                }
                #endregion

                #region nature
                string nature = Nature;
                Thread.Sleep(1000);
                nature = nature.Replace("[^a-zA-Z0-9]", " ");
                string nature_01 = Regex.Replace(nature, @"[^0-9a-zA-Z]+", " ");
                // string nature_01 = nature.Replace("( +.,-&:/)", " ").Trim();
                driver.FindElement(By.XPath(".//*[@id='txtCgDesc_1']")).SendKeys(nature_01);

                //====executed on date
                Executed_On_Date = Executed_On_Date.Substring(0, 11);
                Stringarr = Executed_On_Date.Split('/');
                string dat = Convert.ToString(Stringarr[0]);
                string mon = Convert.ToString(Stringarr[1]).ToUpper();
                string year = Convert.ToString(Stringarr[2]);

                if (mon.Contains("JAN") || mon.Contains("01") || mon.Equals("JANUARY"))
                {
                    mon = "01";
                }
                else if (mon.Contains("FEB") || mon.Contains("02") || mon.Equals("FEBRUARY"))
                {
                    mon = "02";
                }
                else if (mon.Contains("MAR") || mon.Contains("03") || mon.Equals("MARCH"))
                {
                    mon = "03";
                }
                else if (mon.Contains("APR") || mon.Contains("04") || mon.Equals("APRIL"))
                {
                    mon = "04";
                }
                else if (mon.Contains("MAY") || mon.Contains("05") || mon.Equals("MAY"))
                {
                    mon = "05";
                }
                else if (mon.Contains("JUN") || mon.Contains("06") || mon.Equals("JUNE"))
                {
                    mon = "06";
                }
                else if (mon.Contains("JULY") || mon.Contains("07") || mon.Equals("JULY"))
                {
                    mon = "07";
                }
                else if (mon.Contains("AUG") || mon.Contains("08") || mon.Equals("AUGUST"))
                {
                    mon = "08";
                }
                else if (mon.Contains("SEP") || mon.Contains("09") || mon.Equals("SEPTEMBER"))
                {
                    mon = "09";
                }
                else if (mon.Contains("OCT") || mon.Contains("10") || mon.Equals("OCTOBER"))
                {
                    mon = "10";
                }
                else if (mon.Contains("NOV") || mon.Contains("11") || mon.Equals("NOVEMBER"))
                {
                    mon = "11";
                }

                else if (mon.Contains("DEC") || mon.Contains("12") || mon.Equals("DECEMBER"))
                {
                    mon = "12";
                }
                else
                {
                    mon = DateTime.Now.ToString("MM");
                }

                string EXDATE = dat + "/" + mon + "/" + year;
                IWebElement EXDT = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_txtexecutedon']"));
                EXDT.Clear();
                EXDT.SendKeys(EXDATE);
                EXDT.SendKeys(Keys.Tab);

           //     At_Place = At_Place.Replace(" ", "");

                //   string At_Place1 = At_Place.Replace(" ", "");
                //   At_Place1 = At_Place.Substring(0, 3);
               
                //======at place
                if (!string.IsNullOrEmpty(At_Place))
                {
                    if (At_Place.Contains("BANGALORE"))
                    {
                        At_Place = "BENGALURU";
                    }
                    if (At_Place.Contains("BANGLORE"))
                    {
                        At_Place = "BENGALURU";
                    }
                    int cn = At_Place.Length;
                    if (cn <= 3)
                    {
                        IWebElement txtcode = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtCode']"));
                        txtcode.SendKeys(At_Place);
                        txtcode.SendKeys(Keys.Tab);
                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            String alertMessage = driver.SwitchTo().Alert().Text;
                            Console.WriteLine('"' + alertMessage + '"');
                            driver.SwitchTo().Alert().Accept();

                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString().ToUpper();
                                string attp = At_Place.ToUpper();
                                if (attp.Contains(code))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                            /*
                            IWebElement genautcity = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtName']"));
                            genautcity.SendKeys(At_Place);
                            genautcity.SendKeys(Keys.Tab);
                            result = PTEGE.isAlertPresent(driver);
                            if (result == true)
                            {
                                alertMessage = driver.SwitchTo().Alert().Text;
                                Console.WriteLine('"' + alertMessage + '"');
                                driver.SwitchTo().Alert().Accept();
                                IWebElement txtcode1 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtCode']"));
                                txtcode1.SendKeys("BOM");
                                txtcode1.SendKeys(Keys.Tab);
                            }
                            */
                        }
                    }
                    else if (cn > 3)
                    {
                        IWebElement genautcity = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtName']"));
                     //   At_Place = At_Place.Substring(0, 4);
                        genautcity.SendKeys(At_Place);
                        genautcity.SendKeys(Keys.Tab);
                        result = PTEGE.isAlertPresent(driver);
                        if (result == true)
                        {
                            String alertMessage = driver.SwitchTo().Alert().Text;
                            Console.WriteLine('"' + alertMessage + '"');
                            driver.SwitchTo().Alert().Accept();
                         
                            for (int lr = 0; lr < ss; lr++)
                            {
                                string code = dtset_air.Tables["Sheet1"].Rows[lr]["Code"].ToString();
                                string name = dtset_air.Tables["Sheet1"].Rows[lr]["Name"].ToString().ToUpper();
                                string attp = At_Place.ToUpper();
                                if (attp.Contains(name))
                                {
                                    IWebElement airportname = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtName']"));
                                    airportname.SendKeys(name);
                                    airportname.SendKeys(Keys.Tab);
                                    lr = ss;
                                }
                            }
                            /*
                            IWebElement txtcode1 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtCode']"));
                            txtcode1.SendKeys("BOM");
                            txtcode1.SendKeys(Keys.Tab);
                            */
                        }
                    }
                }
                else
                {
                    IWebElement txtcode1 = driver.FindElement(By.XPath(".//*[@id='ctl00_hldPage_GenericAutoFillCity_txtCode']"));
                    txtcode1.SendKeys("BOM");
                    txtcode1.SendKeys(Keys.Tab);
                }

                #endregion

                driver.FindElement(By.XPath("//input[@id='ctl00_hldPage_btnSaveAwb']")).Click();

                try
                {
                    new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementIsVisible((By.Id("ctl00_hldPage_btnCHANo"))));
                    Thread.Sleep(300);
                    string message = driver.FindElement(By.Id("ctl00_hldPage_lblMessage2")).Text;
                    Console.WriteLine(message);
                    if (message.Contains("Air Waybill " + Prefix + "-" + awbno + " saved successfully..."))
                    {
                        strWriter.WriteLine("AWB created for  " + Prefix + "-" + awbno + "");
                        PTEGE.UpdatePingStatus(Prefix, awbno, strWriter);
                        driver.FindElement(By.Id("ctl00_hldPage_btnCHANo")).Click();

                        #region creating daily report

                        string folderpath1 = ConfigurationManager.AppSettings["Folderpath"].ToString();
                        string folderpath = (folderpath1 + DateTime.Now.ToString("dd-MM-yyyy"));
                        string aaa = PTEGE.AutoFolderCreate(folderpath);

                        #endregion
                    }
                    else
                    {
                        PTEGE.changeoperator(Prefix, awbno, oper, strWriter);
                     //   PTEGE.changeoperator(Prefix, awbno, "IntelecS", strWriter);                       
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    PTEGE.changeoperator(Prefix, awbno, oper, strWriter);
                    //   PTEGE.changeoperator(Prefix, awbno, "IntelecS", strWriter);

                    Thread.Sleep(300);
                    driver.FindElement(By.XPath("//a[contains(.,'Log Out')]")).Click();
                    new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.Id("txtUsrName"))));
                    Thread.Sleep(300);
                    driver.Close();
                    driver.Quit();
                }
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.XPath("//a[contains(.,'Home')]"))));
                Thread.Sleep(300);
                driver.FindElement(By.XPath("//a[contains(.,'Log Out')]")).Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(100)).Until(ExpectedConditions.ElementIsVisible((By.Id("txtUsrName"))));
                Thread.Sleep(300);
                driver.Close();
                driver.Quit();
                strWriter.Flush();
                strWriter.Close();
                strWriter.Dispose();
                #endregion
            }
            catch (Exception ex)
            {
                GeneralMethods PTEGE1 = new GeneralMethods();

                PTEGE1.changeoperator(Prefix, awbno, oper, strWriter);
                PTEGE1.UpdatePingStatus(Prefix, awbno, strWriter);
                Console.WriteLine(ex.Message);
                Thread.Sleep(300);              
            }
        }
    }
}
