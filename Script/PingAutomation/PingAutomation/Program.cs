using Microsoft.Office.Interop.Outlook;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using PingAutomation;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace PingAutomation
{
    class Program
    {
        static string ClientName = null,
        Prefix = null,
        AWBNo = null,
        HAWBcount = null,
        Shipmentstatus = null;
        static IWebDriver driver;

        static void Main(string[] args)
        {
            GeneralMethods PTEGE = new GeneralMethods();
       //     string folderpath1 = ConfigurationManager.AppSettings["Folderpath"].ToString();
       //     string folderpath = (folderpath1 + DateTime.Now.ToString("dd-MM-yyyy"));
       //     string aaa = PTEGE.AutoFolderCreate(folderpath);



            string oper = null;
            string Trackernm = ConfigurationManager.AppSettings["Trackername"].ToString().ToUpper();
            if (Trackernm.Contains("NEW"))
            {
                oper = ConfigurationManager.AppSettings["Operatornew"].ToString();
            }
            else if (Trackernm.Contains("OLD"))
            {
                oper = ConfigurationManager.AppSettings["Operatorold"].ToString();
            }

            try
            {
                int MoveFile;
                // create c1 = new create();
                //c1.createawb(@"E:\JITESH\PingAutomation\Output_File\20190114_123638.xls", "MAGNUM CARGO PVT. LTD", "020");
                Console.WriteLine("program is started");
                FileStream files;
                StreamWriter strWriter;
                String FileName = ConfigurationManager.AppSettings["LogfilePath"] + @"\PingAutomation" + DateTime.Now.ToString("dd-MMM-yyyy") + ".txt";
                files = new FileStream(FileName, FileMode.Append, FileAccess.Write, FileShare.None);
                strWriter = new StreamWriter(files);
                strWriter.WriteLine("---------------------------------------------------------------");
                strWriter.WriteLine("Log generated at " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm"));
                strWriter.WriteLine(" ");
                Console.WriteLine("Log");

                //   create cc = new create();
                //    cc.createawb(@"E:\JITESH\PingAutomation\Output_File\20190131_060021.xls", "MAGNUM CARGO PVT. LTD", "020", "50752100", strWriter);

               

                #region datatable of status 2-ACk and hawb count 0
                DataSet TrackerData = new DataSet();
              
                TrackerData = PTEGE.isPresentontracker();
                int o = TrackerData.Tables[0].Rows.Count;
                #endregion
                if (o >= 1)
                {
                    #region get clientdetails from tracker

                    for (int w = 0; w < o; w++)
                    {
                        Console.WriteLine("clientdetails");
                        string  ClientName1 = TrackerData.Tables["Table1"].Rows[w]["Client_Name"].ToString();

                        int clen=ClientName1.Length;
                        if (clen < 15)
                        {
                            int sss = clen - 1;
                            ClientName = ClientName1.Substring(0, clen);
                        }
                        else
                        {
                            ClientName = ClientName1.Substring(0, 15);
                        }
                        Prefix = TrackerData.Tables["Table1"].Rows[w]["Pfx"].ToString();
                        AWBNo = TrackerData.Tables["Table1"].Rows[w]["AWB_No"].ToString();
                        HAWBcount = TrackerData.Tables["Table1"].Rows[w]["Hawb Count"].ToString();
                        Shipmentstatus = TrackerData.Tables["Table1"].Rows[w]["Shipment_Status"].ToString();
                        Console.WriteLine(ClientName + " " + Prefix + " " + AWBNo);
                        string a = AWBNo.Substring(0, 4);
                        string b = AWBNo.Substring(4, 4);
                        string spaceAWB = a + " " + b;

                        DataSet dd = new DataSet();
                        string excelpath = ConfigurationManager.AppSettings["RuleExcel_Ping"].ToString();
                        dd = PTEGE.ImportexcelData(excelpath, ClientName);
                        #endregion
                        int v = dd.Tables[0].Rows.Count;
                        Thread.Sleep(300);
                        for (int h = 0; h < v; h++)
                        {
                            Thread.Sleep(300);
                            string Prefixexcl = dd.Tables["Sheet1"].Rows[h][1].ToString();

                            if (Prefixexcl.Contains(Prefix))
                            {
                                Thread.Sleep(300);
                                string Rule_Nm = dd.Tables["Sheet1"].Rows[h][2].ToString();
                                if (!Rule_Nm.Contains("Not_Done"))
                                {                                 

                                    #region pst
                                    string fpath = null;
                                    try
                                    {
                                        IEnumerable<MailItem> mailItems = readPst(ConfigurationManager.AppSettings["PSTFilePath"].ToString(), ConfigurationManager.AppSettings["PSTFolder"].ToString());
                                        foreach (MailItem mailItem in mailItems)
                                        {
                                            string ReceivedTime = string.Empty, Sender = string.Empty, MailBody = string.Empty;
                                            string Subject = string.Empty;

                                            //Sender = mailItem.SenderEmailAddress;
                                            Sender = mailItem.SenderName;
                                            Subject = mailItem.Subject;
                                            {
                                                List<string> tempFileName = new List<string>();
                                                List<string> tempPDFFileName = new List<string>();

                                                //  string temp = mailItem.Attachments[1].FileName;

                                                String FolderName = NewFolderName();
                                                int k = 1;

                                                int j = mailItem.Attachments.Count;


                                                if (mailItem.Attachments.Count == 0)
                                                {
                                                    Application app1 = new Application();
                                                    NameSpace outlookNs1 = app1.GetNamespace("MAPI");
                                                    MAPIFolder MoveToFolder1 = outlookNs1.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderDeletedItems);
                                                    mailItem.Move(MoveToFolder1);
                                                    strWriter.WriteLine(mailItem.Subject + " Have Moved to Deleted folder successfully as attachment is not present");
                                                    strWriter.WriteLine(" ");
                                                    strWriter.WriteLine("-------------------------------------------------------");
                                                    strWriter.WriteLine(" ");
                                                    MoveFile = 0;
                                                    //  break;
                                                }


                                                Console.WriteLine("mail attachment");
                                                if (mailItem.Attachments.Count >= 1)
                                                {                                                  
                                                    if (!string.IsNullOrEmpty(Subject))
                                                    {                                                       
                                                        if (Subject.Contains(AWBNo) || Subject.Contains(spaceAWB))
                                                        {
                                                            DataSet ex = new DataSet();
                                                            excelpath = ConfigurationManager.AppSettings["RuleExcel_Ping"].ToString();
                                                            ex = PTEGE.ImportexcelData(excelpath, ClientName);
                                                            int p = ex.Tables[0].Rows.Count;

                                                            for (int x = 0; x < p; x++)
                                                            {
                                                                Thread.Sleep(300);
                                                                Console.WriteLine("RULECOUNT = "+x);
                                                                string Prefixexcel = ex.Tables["Sheet1"].Rows[x][1].ToString();
                                                                Thread.Sleep(300);
                                                                if (Prefix.Equals(Prefixexcel))
                                                                {
                                                                    Thread.Sleep(300);
                                                                    string Rule_Name = ex.Tables["Sheet1"].Rows[x][2].ToString();
                                                                    if (!Rule_Name.Contains("Not_Done"))
                                                                    {
                                                                        Console.WriteLine("Rule present");
                                                                        for (int g = 1; g <= mailItem.Attachments.Count; g++)
                                                                        { 
                                                                            if (mailItem.Attachments[g].FileName.Contains(".pdf") || mailItem.Attachments[g].FileName.Contains(".PDF"))
                                                                            {
                                                                                Console.WriteLine("attachment downloaded");
                                                                                fpath = ConfigurationManager.AppSettings["PDFFolderPath"].ToString();
                                                                                mailItem.Attachments[g].SaveAsFile(fpath + mailItem.Attachments[g].FileName);
                                                                                string filename = fpath + mailItem.Attachments[g].FileName;
                                                                                Console.WriteLine(filename);
                                                                                string newfilename;
                                                                                string newpath;

                                                                                bool fHasSpace = filename.Contains(" ");
                                                                                if (fHasSpace == true)
                                                                                {
                                                                                    newpath = filename.Replace(" ", "");
                                                                                    Console.WriteLine(newpath);
                                                                                    //  newfilename = Prefix + "-" + AWBNo;

                                                                                    if (File.Exists(newpath))
                                                                                    {
                                                                                        File.Delete(newpath);
                                                                                    }
                                                                                    System.IO.File.Move(filename, newpath);
                                                                                }
                                                                                else
                                                                                {
                                                                                    newpath = filename;
                                                                                    Console.WriteLine(newpath);
                                                                                }

                                                                                Console.WriteLine("Rule creation");
                                                                                string rulename = Rule_Name + "_" + Prefixexcel;
                                                                                String targetfile = DateTime.Now.ToString("yyyyMMdd_hhmmss");
                                                                                String targetfile1 = targetfile + ".xls";
                                                                                string targetfolderpath = ConfigurationManager.AppSettings["Output_File_Path"].ToString();
                                                                                string outputfile = targetfolderpath + targetfile1;
                                                                                Process process = new Process();
                                                                                process.StartInfo.FileName = "cmd.exe";
                                                                                process.StartInfo.CreateNoWindow = true;
                                                                                process.StartInfo.RedirectStandardInput = true;
                                                                                process.StartInfo.RedirectStandardOutput = true;
                                                                                process.StartInfo.UseShellExecute = false;
                                                                                process.Start();
                                                                                Thread.Sleep(2000);
                                                                                process.StandardInput.WriteLine("PDECMD -R\"" + rulename + "\" -F\"" + newpath + "\" -O\"" + outputfile);
                                                                                Thread.Sleep(8000);
                                                                                string xyz = process.StandardOutput.ReadLine();
                                                                                process.StandardInput.Flush();
                                                                                process.StandardInput.Close();
                                                                                process.WaitForExit();
                                                                                Thread.Sleep(2000);
                                                                                Console.WriteLine(process.StandardOutput.ReadToEnd());
                                                                                string xyz1 = process.StandardOutput.ReadLine();
                                                                                string outputReader = process.StandardOutput.ReadToEnd();
                                                                                PTEGE.changeoperator(Prefix, AWBNo, "Automation", strWriter);
                                                                                driver = new ChromeDriver();
                                                                                Console.WriteLine("rule created");
                                                                                Thread.Sleep(2000);
                                                                                create c = new create();
                                                                                c.createawb(outputfile, ClientName, Prefix, AWBNo, strWriter,driver);
                                                                                w = o;                                                                              
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        Console.WriteLine(ex.Message);
                                        PTEGE.changeoperator(Prefix, AWBNo, oper, strWriter);
                                        PTEGE.UpdatePingStatus(Prefix, AWBNo, strWriter);                                      
                                        //   PTEGE.changeoperator(Prefix, AWBNo, "IntelecS", strWriter);
                                        driver.Close();
                                        driver.Quit();
                                        break;
                                    }
                                    #endregion
                                }
                            }
                        }


                    }
                }

                driver.Close();
                driver.Quit();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
              
            }
        }

        #region splitstring
        static IEnumerable<string> Split(string str, int chunkSize)
        {
            return Enumerable.Range(0, str.Length / chunkSize)
                .Select(i => str.Substring(i * chunkSize, chunkSize));
        }
        #endregion

        #region PST read

        private static IEnumerable<MailItem> readPst(string pstFilePath, string pstName) // Will return the list of mail items present in Inbox....
        {
            List<MailItem> mailItems = new List<MailItem>();
            Application app = new Application();
            NameSpace outlookNs = app.GetNamespace("MAPI");
            // Add PST file (Outlook Data File) to Default Profile
            outlookNs.AddStore(pstFilePath);
            int COUNT = outlookNs.Folders.Count;

            MAPIFolder inboxFolder = outlookNs.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            Items mailItems1 = inboxFolder.Items;
            foreach (MAPIFolder folder in outlookNs.Folders)
            {              
                if (folder.FolderPath == "\\\\" + ConfigurationManager.AppSettings["MailFolder"].ToString())
                {                   
                    foreach (MAPIFolder subFolder in folder.Folders)
                    {                        
                       // if (subFolder.FolderPath == "\\\\" + ConfigurationManager.AppSettings["MailFolder"].ToString() + "\\Inbox")
                            if (subFolder.FolderPath == "\\\\" + ConfigurationManager.AppSettings["MailFolder"].ToString() + "\\Drafts")
                            {
                            // foreach (MAPIFolder InboxSubFolder in subFolder.Folders)
                          //  foreach (MAPIFolder DRAFTSFolder in subFolder.Folders)
                           {                                
                               // if (InboxSubFolder.FolderPath == "\\\\" + ConfigurationManager.AppSettings["MailFolder"].ToString() + "\\Inbox\\PING_TRACKER")
                                {
                                    //foreach (Microsoft.Office.Interop.Outlook.MailItem item in InboxSubFolder.Items)
                                        foreach (Microsoft.Office.Interop.Outlook.MailItem item in subFolder.Items)
                                        {
                                        if (item is MailItem)
                                        {
                                            
                                            mailItems.Add(item);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return mailItems;
        }

        #endregion

        #region newfolder
        static string NewFolderName()
        {
            // Eliminating following Special Characters from subject ............. " * / : < > ? \ | 
            string f = DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss");

            return f;
        }
        #endregion
    }
}
