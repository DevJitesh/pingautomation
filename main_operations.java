package Create_AirwayBill;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

public class main_operations
{
	static Boolean result=false; 
	static Boolean result1,result2,result3,result4;
	
	
	public static void main(String[] args) throws Exception
	{
		 	
		
		String test="KALE LOGISTICS";
		System.out.println(test.charAt(0));
			// Workbook wb1 = new XSSFWorkbook();
			
//			 org.apache.poi.ss.usermodel.Sheet sheet = wb1.createSheet("AWB");
//			 Row row1 = sheet.createRow(0);
//				row1.createCell(0).setCellValue("Air Waybill Number");
//				row1.createCell(1).setCellValue("Duration");	
			int Srno=1;
			
			String AWB_Number=null;
			String Agent_Name=null;
			String  StartTime=null;
			String EndTime=null;
			String Duration=null;
			String Status=null;
			String Reason=null;
			String ScreenShotPath=null;
			int rono=1;
			String filename=null;
			String excelsave_path;
			String Reportsave_path;
			NumberFormat formatter = new DecimalFormat("#0.00");
			long second = 1000l;
			long minute = 60l * second;
			long hour = 60l * minute;
			String starttime1=null;
			String endtime1=null;
			long diff;
 			 File ocr_config_file = new File("ocr_config_file.properties");
			 InputStream inputStream = new FileInputStream(ocr_config_file);
			  
			 Properties props = new Properties();
			 
			 props.load(inputStream);
			 
			 String Report_folder_path=props.getProperty("Report_folder_path"); 
			
				String Reportname=new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
				
//			      FileOutputStream fileOut = new FileOutputStream(Report_folder_path+"REPORT_"+Reportname+".xls");
//======================================================================================================================

				String s = new SimpleDateFormat("dd.MM.yyyy").format(Calendar.getInstance().getTime());

				String folderpath = props.getProperty("Report_folder_path") + s;
				 String Daily_Report=props.getProperty("Daily_Report")+s;
			//	Webdriver_Components webcomp=new Webdriver_Components();
				write_excel wx=new write_excel();
			 excelsave_path= wx.foldercreate(folderpath);
			 

			 Daily_Report DR=new Daily_Report();
			Reportsave_path= DR.foldercreate_D(Daily_Report);
			
		//	DR.read_write_D("jitesh", "kale",1, Reportsave_path);
    	
//=======================================================================================================================			 
			
			 
	 try
		{	 
		 String chromedriver_path=props.getProperty("chromedriver_path");
			System.setProperty("webdriver.chrome.driver",chromedriver_path);
		WebDriver driver= new ChromeDriver();	 
		driver.manage().window().maximize();
		Wait wait = new FluentWait(driver);
		String TestData_path=props.getProperty("TestData_path");
		String RuleData_path=props.getProperty("RuleData_path");
		excel_operations eat1=new excel_operations(TestData_path);
		excel_operations eat3=new excel_operations(RuleData_path);
		
		String LIVEurl=eat1.getCellData("login(LIVE)","URL",1);
		String UATurl=eat1.getCellData("login(UAT)","URL",1);
		String Sc;
		Sc=eat1.screenShot(driver);
		/*UAT*/
		
//		driver.get(UATurl);
//		String username=eat1.getCellData("login(UAT)","UN",1);
//		String password=eat1.getCellData("login(UAT)","PW",1);
		

		/*LIVE*/
		
		driver.get(LIVEurl);
		String username=eat1.getCellData("login(LIVE)","UN",1);
		String password=eat1.getCellData("login(LIVE)","PW",1);

		driver.findElement(By.id("txtUsrName")).sendKeys(username);
		driver.findElement(By.id("txtPswd")).sendKeys(password);
		driver.findElement(By.id("btnLogin")).click();

		
//==============================================================================================================================================
		
		FileInputStream fis=new FileInputStream(TestData_path);
		@SuppressWarnings("resource")
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet AWBdetails=wb.getSheetAt(1);
		int no=AWBdetails.getLastRowNum();
//		for(int 1=1;1<=no;1++)
//		{
		
		driver.findElement(By.xpath(".//*[@id='liquid-bd']/div[1]/div[1]/ul/li[6]")).click();     //Click AWB Track Ping
//		String actualnumber01=eat1.getCellData("AWBdetails", "actual01",1);
//		String actualnumber02=eat1.getCellData("AWBdetails", "actual02",1);

		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
		Select yts=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
		yts.selectByVisibleText("Yet To Start");                                                   // Select Yet To Start
		
//==========================================================================================================================================
											 
		result=eat1.isElementPresent(driver, "//td[contains(.,'No records found')]");
		if(result==false)
		{
		result=eat1.isElementPresent(driver, ".//*[@id='ctl00_hldPage_grdAWBTrack']/tbody/tr[7]/td/table/tbody/tr[1]/td[1]");
		
		if(result==true)
		{
			WebElement pagecount=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack']/tbody/tr[7]/td/table/tbody/tr[1]/td[1]"));
			String countstring=pagecount.getText();
			String[] splited = countstring.split(" ");
		
			String pagenumber=splited[3];
			int Pageno = Integer.parseInt(pagenumber);		
			for(int p=1;p<=Pageno;p++)
		{
			if(p>1)
			{
			WebElement nextpage=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl08_BtnNext']"));
			Thread.sleep(2000);
         	nextpage.click();
         	Thread.sleep(5000);
			}
		
		
		List<WebElement> all = driver.findElements(By.xpath("//input[contains(@id,'txtOtherComments')]"));
        String[] allText = new String[all.size()];
        int i1 = 2; 

        for (WebElement element : all)
        {	 
//        	result4=eat1.isElementPresent(driver, ".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgAssgnDateTime']");
//        	if(result4==false) 
//        	{
        	result=eat1.isElementPresent(driver, ".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgFinalUpload']");
        	
        	if(result==false)
        	{
        	WebElement housecount=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtHAWBCount']"));
        	 String count = housecount.getAttribute("value");
        
        	 if(count.contains("0"))
        	 {
        		 result1=eat1.isElementPresent(driver, ".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgDraftUpload']");
       
            if (result1== true)
			{
            	result2=eat1.isElementPresent(driver,".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgAWBCreation']");
            	
            if(result2==false)
            {
            	Actions draftup = new Actions(driver);
            	WebElement imgDraftUpload = driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgDraftUpload']"));
            	Thread.sleep(2000);
            	draftup.clickAndHold(imgDraftUpload).perform();
            	String DraftUploadDate_Time = imgDraftUpload.getAttribute("title");
            	
            	String[] arrOfStr1 = DraftUploadDate_Time.split(" ");
				String DraftUpload_time = arrOfStr1[1];
            	
            	starttime1 = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
  		        StartTime=starttime1;
  		          		        
            	WebElement AWBno=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_lblawbno']"));
            	String AWBNumber=AWBno.getText();
            	
            	String [] WholeAWB  = AWBNumber.split("-");
            	String onlyPrefix = WholeAWB [0];
            	String onlyAWB=WholeAWB [1];
            	AWB_Number=onlyAWB;
            	     	
               	onlyPrefix = onlyPrefix.replaceAll(" ", "");
        		    		
            	WebElement OrgName=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack']/tbody/tr["+i1+"]/td[17]"));
            	String OrganizationName=OrgName.getText();
            	System.out.println(OrganizationName);
            	Agent_Name=OrganizationName;
            	 
                 int r=0; String RuleName=null;
                for(r=1;r<=59;r++)
                {
             	 String OrganiseName=eat3.getCellData("RuleMaster",r,"Agent_Name");	
             	 
             	 if(OrganiseName.equals(OrganizationName))
             	 {
             		RuleName=eat3.getCellData("RuleMaster",r,"Rule_Name");	
             	 }
                }
                if(!RuleName.equals("Scan") && !RuleName.equals("Not_Done") )
                {	
                	String UpdateRuleName=RuleName+"_"+onlyPrefix;
                	//String UpdateRuleName="SIAM_098";
                	
//=====================================================================================================================================================================
                	
                	
                	
                	
                	
                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_drpLinks']")).click();
            		Select status=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_drpLinks']")));
            		status.selectByVisibleText("Assigned");
            		Thread.sleep(5000);
            		driver.switchTo().alert().accept();
            		Thread.sleep(5000);

            	    if(username.contains("tiffaawb"))
            	    {
            	    	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtOperator']")).sendKeys("testair");
                    	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtAssignedTo']")).sendKeys("testair");
            	    }
            	    else
            	    {
                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtOperator']")).sendKeys("Automation");
                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtAssignedTo']")).sendKeys("Automation");
            	    }
                	
               	
                    		
            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_lnkSave']")).click();
            		Thread.sleep(10000);
            		new WebDriverWait(driver, 250).until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")));
            		//webcomp.elementToBeClickable(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']"));
            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
            		Thread.sleep(7000);
            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
            		
            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
            		
            		
            		//Thread.sleep(2000);

                	
                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkViewLatest']")).click();
                    Thread.sleep(10000);
            		driver.switchTo().frame("Ifram1");
            		Thread.sleep(10000);
            		//new WebDriverWait(driver, 250).until(ExpectedConditions.visibilityOfElementLocated(By.linkText("view")));
            		driver.findElement(By.linkText("View")).click();
            		Thread.sleep(10000);
                
//=============================================================================================================================================
		/**
		   * Files extraction.
		   */
        String Chrome_Downloads_path=props.getProperty("Chrome_Downloads_path");  
        String Source_folder_path=props.getProperty("Source_folder_path"); 
        String Target_folder_path=props.getProperty("Target_folder_path"); 
        String Proccessed_folder_path=props.getProperty("Proccessed_folder_path"); 
        String Proccess_fail_folder_path=props.getProperty("Proccess_fail_folder_path");
		final File folder = new File(Chrome_Downloads_path);
     String fileName = "null";
     for (final File fileEntry : folder.listFiles())
    {
         	  listFilesForFolder(fileEntry);
             fileName=fileEntry.getName();
//===============================================================================================================================================
             Path temp = Files.move
           	        (Paths.get(Chrome_Downloads_path+fileName), 
           	        Paths.get(Source_folder_path+fileName));
           	 
           	        if(temp != null)
           	        {
           	            
           	        }
           	        else
           	        {
           	            System.out.println("Failed to move the file");
           	        }
             
//===============================================================================================================================================              
             File file = new File(Source_folder_path+fileName);
             String renamefile=new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
             String renamefile1=renamefile+".pdf";
             File input_file = new File(Source_folder_path+renamefile1);
             if(file.renameTo(input_file))
             {
                 
             }
             else
             {
                 System.out.println("File rename failed");
             }
//===============================================================================================================================================              
             
             String targetfile=new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
             String targetfile1=targetfile+".xls";
             String TARGET_FILE=Target_folder_path+targetfile1;
//===========================================================================================================================
    try
     { 
     
   	//  Runtime.getRuntime().exec("cmd /c start PDECMD -R\"OCR_PDF_RULE\" -F\""+newFile+"\" -O\"D:\\VIVEK\\TIFFA\\OCR\\Target\\pdfdata -TXLS -PA.xls");
   	 // Runtime.getRuntime().exec("cmd /c start PDECMD -R\"SIAM_KARGO_LOGISTICS\" -F\""+newFile+"\" -O\"D:\\VIVEK\\TIFFA\\OCR\\Target\\"+targetfile1);
    	 Runtime.getRuntime().exec("cmd /c start PDECMD -R\""+UpdateRuleName+"\" -F\""+input_file+"\" -O\""+TARGET_FILE);
   	  Thread.sleep(8000);
     }
     
     catch (Exception e)
     {
         System.out.println("Incorrect action perform");
         e.printStackTrace();
     }
//========================================================================================================================================================	
    File f = new File(TARGET_FILE);
    if(f.exists() && !f.isDirectory()) 
    { 
        
    
//========================================================================================================================================================    
    Path temp1 = Files.move
 	        (Paths.get(Source_folder_path+renamefile1), 
 	        Paths.get(Proccessed_folder_path+renamefile1));
 	 
 	        if(temp1 != null)
 	        {
 	           
 	        }
 	        else
 	        {
 	            System.out.println("Failed to move the file");
 	        }
//=========================================================================================================================================================      
 	       excel_operations eat2=new excel_operations(Target_folder_path+targetfile1);
			driver.switchTo().defaultContent();
			Thread.sleep(5000);
			WebElement close=driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[6]/div[11]/button"));
			Thread.sleep(10000);
			close.click();

			
			 ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());
			 
//			 driver.close();
//			 Thread.sleep(3000);
		
		   driver.switchTo().window(tabs2.get(1));
		
			    Thread.sleep(10000);
		  
		long start = System.currentTimeMillis();
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBPrefix']")).sendKeys(onlyPrefix);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(String.valueOf(onlyAWB));
		String ShipName=eat2.getCellData("Sheet1", "Shipper_Name",1);
		if(!ShipName.equals("data not found"))
		{
//======================================================================================================================================	
		driver.findElement(By.xpath(".//*[@id='btnshipperadd']")).click();
		
		
		ShipName = ShipName.replaceAll("[^a-zA-Z0-9]", " ");
		String ShipName_01=ShipName.replaceAll("( +)"," ").trim();
		
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCompanyName']")).sendKeys(ShipName_01);
		
//		int ShipName_length=ShipName_01.length();
//		if(ShipName_length >35 )
//		{
//		String limit_ShipName = ShipName_01.substring(0, 35);
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCompanyName']")).sendKeys(limit_ShipName);
//		}
//		else
//		{
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCompanyName']")).sendKeys(ShipName);
//		}
//=======================================================================================================================================				
		String Remain_Addline="";String Concate_add2_add1=null;
		String Addline_01=eat2.getCellData("Sheet1", "Address_Line01",1);
		Thread.sleep(1000);
		Addline_01 = Addline_01.replaceAll("[^a-zA-Z0-9]", " ");
		String MAddline_01=Addline_01.replaceAll("( +)"," ").trim();
		
		String Addline_02=eat2.getCellData("Sheet1", "Address_Line02",1);
		String Concate_AddLine02;
		Thread.sleep(1000);
		Addline_02 = Addline_02.replaceAll("[^a-zA-Z0-9]", " ");
		String MAddline_02=Addline_02.replaceAll("( +)"," ").trim();
		
		
		if(MAddline_01=="data not found" || MAddline_01=="" || MAddline_01==" ")
		{
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 1");	
		}
		else
		{
		int Addline_length =MAddline_01.length();
		
		if(Addline_length >35)
		{
				String limit_Addline = MAddline_01.substring(0, 35);
				 Remain_Addline=MAddline_01.substring(36, Addline_length);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys(limit_Addline);
		}
		else
		{
			Thread.sleep(1000);
			
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys(MAddline_01);
		}
		}
//===============================================================================================================================================		
		
		if(MAddline_02=="data not found" || MAddline_02==" " || MAddline_02=="")
		{
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 2");
		}
		else
		{
			if(Remain_Addline=="")
			{
				Concate_AddLine02=MAddline_02;
			}
			else
			{
				Concate_AddLine02=Remain_Addline+MAddline_02;
			}
		
		int Addline_length_02 =Concate_AddLine02.length();
		if(Addline_length_02 >35)
		{
				
		String limit_Addline_02=Concate_AddLine02.substring(0,35);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine2']")).sendKeys(limit_Addline_02);
		}
		else
		{
			Thread.sleep(1000);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine2']")).sendKeys(
);
		}
		}
//============================================================================================================================================		
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCountry_txtCode']")).sendKeys("IN");
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillState_txtCode']")).sendKeys("MH");
//
//		Thread.sleep(1000);
////		String CityName=eat2.getCellData("Sheet1", "City",2);
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCity_txtCode']")).sendKeys("BOM");
//		Thread.sleep(1000);
//
////		String PinCodeNo=eat2.getCellData("Sheet1", "PinCode",1);
//		WebElement we=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtPinCode']"));
//		we.sendKeys("400610");
//		we.click();
//		
//
//       
//		Thread.sleep(1000);
//		we.sendKeys(Keys.TAB,Keys.ENTER);
//===============================================================================================================================================
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCountry_txtCode']")).sendKeys("TH");
		

		Thread.sleep(1000);
//		String CityName=eat2.getCellData("Sheet1", "City",2);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCity_txtCode']")).sendKeys("BKK");
		Thread.sleep(1000);

//		String PinCodeNo=eat2.getCellData("Sheet1", "PinCode",1);
		WebElement we=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtPinCode']"));
		we.sendKeys("XXXXXX");
		we.click();
		

       
		Thread.sleep(1000);
		we.sendKeys(Keys.TAB,Keys.ENTER);
//===============================================================================================================================================		
		
		driver.findElement(By.xpath(".//*[@id='imgbtnConsignee']")).click();
		
		String ConsigName=eat2.getCellData("Sheet1", "Consignee_Name",1);
		
		
		ConsigName = ConsigName.replaceAll("[^a-zA-Z0-9]", " ");
		String ConsigName_01=ConsigName.replaceAll("( +)"," ").trim();
		int ConsigName_length=ConsigName_01.length();
		if(ConsigName_length >35)
		{
		String limit_ConsigName = ConsigName_01.substring(0, 35);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtConName']")).sendKeys(limit_ConsigName);
		}
		else
		{
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtConName']")).sendKeys(ConsigName);	
		}
//=====================================================================================================================================		
		String Remain_Addline_02="";
		String CAddline_01=eat2.getCellData("Sheet1", "Address_Line01_c",1);
		
		Thread.sleep(1000);
		CAddline_01 = CAddline_01.replaceAll("[^a-zA-Z0-9]", " ");
		String CMAddline_01=CAddline_01.replaceAll("( +)"," ").trim();
		if(CMAddline_01=="data not found" || CMAddline_01=="" || CMAddline_01==" ")
		{
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 1");	
		}
		else
		{
		int CAddline1_length=CMAddline_01.length();
		if(CAddline1_length >35)
		{
		String limit_Addline01 = CMAddline_01.substring(0, 35);
		Remain_Addline_02=CMAddline_01.substring(36, CAddline1_length);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine1']")).sendKeys(limit_Addline01);
		}
		else
		{
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine1']")).sendKeys(CMAddline_01);
		}
		}
//=========================================================================================================================
		
		String CAddline_02=eat2.getCellData("Sheet1", "Address_Line02_c",1);
		String Concate_CAddLine02;
		Thread.sleep(1000);
		CAddline_02 = CAddline_02.replaceAll("[^a-zA-Z0-9]", " ");
		String CMAddline_02=CAddline_02.replaceAll("( +)"," ").trim();
		if(CMAddline_02=="data not found" || CMAddline_02=="" || CMAddline_02==" ")
		{
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 2");	
		}
		else
		{
			if(Remain_Addline_02=="")
			{
				Concate_CAddLine02=CMAddline_02;
			}
			else
			{
				Concate_CAddLine02=Remain_Addline_02+CMAddline_02;
			}
			
		int CAddline2_length=Concate_CAddLine02.length();
		if(CAddline2_length >35)
		{
		String limit_Addline02 = Concate_CAddLine02.substring(0, 35);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine2']")).sendKeys(limit_Addline02);
		}
		else
		{
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine2']")).sendKeys(Concate_CAddLine02);
		}
		}		
//=========================================================================================================================		
//		//String CountryName1=eat1.getCellData("AWBdetails", "Country_c",1);
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCountry_txtName']")).sendKeys("United Arab Emirates");
//		
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillState_txtCode']")).sendKeys("DX");
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCity_txtCode']")).sendKeys("DXB");
//		
//		//String CityName1=eat2.getCellData("Sheet1", "City_c",2);
//		//driver.findElement(By.id("ctl00_hldPage_uclConsAddressInfo_txtOtherCity")).sendKeys("MUMBAI");
//		
//		
//		//String PinCode1=eat2.getCellData("Sheet1", "PinCode_c",2);
//		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtPinCode']")).sendKeys("000000");
//		Thread.sleep(1000);
//		
//		WebElement wf=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclAlsoNotifyAddressInfo_txtPinCode']"));
//		wf.click();
//		wf.sendKeys(Keys.TAB,Keys.ENTER);
//		Thread.sleep(1000);
//==============================================================================================================================================
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCountry_txtName']")).sendKeys("INDIA");
		
		
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCity_txtCode']")).sendKeys("BOM");
		
		//String CityName1=eat2.getCellData("Sheet1", "City_c",2);
		//driver.findElement(By.id("ctl00_hldPage_uclConsAddressInfo_txtOtherCity")).sendKeys("MUMBAI");
		
		
		//String PinCode1=eat2.getCellData("Sheet1", "PinCode_c",2);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtPinCode']")).sendKeys("XXXXXX");
		Thread.sleep(1000);
		
		WebElement wf=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclAlsoNotifyAddressInfo_txtPinCode']"));
		wf.click();
		wf.sendKeys(Keys.TAB,Keys.ENTER);
		Thread.sleep(1000);
//===============================================================================================================================================
//Issuing Carrier's Agent Name and City
		String IssuAgent=eat2.getCellData("Sheet1", "Agent",1);
		
		WebElement agnt=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAgentSelect']"));
//		if(IssuAgent.equals("data not found") && IssuAgent.equals("  "))
		//{
				if(OrganizationName.equalsIgnoreCase("TIFFA AWB Service"))
				{
					IssuAgent="KALE LOGISTICS";
				}
				else
				{
				IssuAgent=OrganizationName;
				}
				//IssuAgent="SIAM KARGO LOGISTICS CO LTD";
				//IssuAgent="EAGLES AIR AND SEA (THAILAND) CO., LTD";
				
	//	  }
	//	String AgentConcate= IssuAgent.substring(0, IssuAgent.indexOf("."));
		
		agnt.sendKeys(IssuAgent);
		Thread.sleep(2000);
		agnt.sendKeys(Keys.CONTROL+"A");
		Thread.sleep(2000);
		agnt.sendKeys(Keys.ARROW_DOWN);
		Thread.sleep(1000);
		agnt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		
		String AccountInformatiom=eat2.getCellData("Sheet1", "Accounting_Information",1);
		AccountInformatiom = AccountInformatiom.replaceAll("[^a-zA-Z0-9]", " ");
		String AccountInformatiom_01=AccountInformatiom.replaceAll("( +)"," ").trim();
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_accountinginfo']")).sendKeys(AccountInformatiom_01);
//================================================================================================================================
		//flight details
		String Carrier_code_01=null,Airline_Date_01,Airline_Date_001,Flight_Number_01=null,Current_Date,Flight_date01=null,Carrier_code_02=null,Airline_Date_02=null,Airline_Date_002,Flight_Number_02=null,Flight_date02=null,Carrier_code_03=null,Airline_Date_03=null,Airline_Date_003,Flight_Number_03=null,Flight_date03=null;
		 String [] Array_flightdetails_01,Array_flightdetails_02,Array_flightdetails_03;
				 Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
				 
				 String OriPort=eat2.getCellData("Sheet1", "Origin_Port",1);
				 
				 System.out.println(OriPort.charAt(0));
				 
				 String OriPort_01=OriPort.replaceAll("( +)"," ").trim();
				 if(OriPort_01=="" ||  OriPort=="data not found")
				 {
					 OriPort_01="data not found"; 
				 }
				 
				 if(OriPort_01.equals("SUVARNABHUMI AIRPORT , THAILAND"))
				 {
					 OriPort_01="BKK";
				 }
				 else if(OriPort_01.equals("BANGKOK/THAILAND"))
				 {
					 OriPort_01="BKK";
				 }
				 else if(OriPort_01.equals("BANGKOK,THAILAND"))
				 {
					 OriPort_01="BKK";
				 }
				 else if(OriPort_01.equals("SUVARNABHUMI AIRPORT"))
				 {
					 OriPort_01="BKK";
				 }
				 else if(OriPort_01.equals("DON MUEANG AIRPORT, THAILAND"))
				 {
					 OriPort_01="DMK";
				 }
				 else if(OriPort_01.equals("DON MUEANG AIRPORT"))
				 {
					 OriPort_01="DMK";
				 }
				 
				String Via1=eat2.getCellData("Sheet1", "Via01",1);
				String Via1_01=Via1.replaceAll("( +)"," ").trim();
				if(Via1_01=="" || Via1=="data not found")
				 {
					Via1_01="data not found"; 
				 }
				
				String Via2=eat2.getCellData("Sheet1", "Via02",1);
				String Via2_01=Via2.replaceAll("( +)"," ").trim();
				
				if(Via2_01=="" || Via2=="data not found")
				 {
					Via2_01="data not found"; 
				 }
				
				String DestPort=eat2.getCellData("Sheet1", "Destn_Port",1);
				String DestPort_01=DestPort.replaceAll("( +)"," ").trim();
				if(DestPort_01==("") || DestPort=="data not found")
				 {
					DestPort_01="data not found"; 
				 }
				
		
				String Flight_Details_01=eat2.getCellData("Sheet1", "Flight_Details_01",1);
				String Flight_Details_02=eat2.getCellData("Sheet1", "Flight_Details_02",1);
				String Flight_Details_03=eat2.getCellData("Sheet1", "Flight_Details_03",1);
				
				if(IssuAgent=="KALE LOGISTICS")
				{
					Flight_Details_01="9W";
					Flight_Details_02="9W1922/26";
					Flight_Details_03="data not found";
				}
				int lenght_fli_det_01=Flight_Details_01.length();
				if(lenght_fli_det_01<=4)
				{
					Flight_Details_01="data not found";
				}
				
				
				if(DestPort=="data not found")
				{
					Flight_Details_01=Flight_Details_02;
					Flight_Details_02=Flight_Details_03;
				}
				else if(Via2=="data not found")
				{
					Flight_Details_01=Flight_Details_02;
				}
				
				if(Flight_Details_03 !="data not found")
				{				
					Array_flightdetails_01  = Flight_Details_01.split("/");
					 Carrier_code_01=Flight_Details_01.substring(0,2);
					 Airline_Date_001=Array_flightdetails_01[1];
					 Airline_Date_01=Airline_Date_001.substring(0,2);
					 Flight_Number_01=Array_flightdetails_01[0].substring(2);
					 Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
					 Flight_date01=Airline_Date_01+Current_Date;
				 
					 Array_flightdetails_02  = Flight_Details_02.split("/");
					 Carrier_code_02=Flight_Details_02.substring(0,2);
					 Airline_Date_002=Array_flightdetails_02[1];
					 Airline_Date_02=Airline_Date_002.substring(0,2);
					 Flight_Number_02=Array_flightdetails_02[0].substring(2);
					 Flight_date02=Airline_Date_02+Current_Date;
				
					 Array_flightdetails_03  = Flight_Details_03.split("/");
					 Carrier_code_03=Flight_Details_03.substring(0,2);
					 Airline_Date_003=Array_flightdetails_03[1];
					 Airline_Date_03=Airline_Date_003.substring(0,2);
					 Flight_Number_03=Array_flightdetails_03[0].substring(2);
					 Flight_date03=Airline_Date_03+Current_Date;
				}

				else if(Flight_Details_02 !="data not found")
				{
						Array_flightdetails_01  = Flight_Details_01.split("/");
						Carrier_code_01=Flight_Details_01.substring(0,2);
						Airline_Date_001=Array_flightdetails_01[1];
						Airline_Date_01=Airline_Date_001.substring(0,2);
						Flight_Number_01=Array_flightdetails_01[0].substring(2);
						Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
						Flight_date01=Airline_Date_01+Current_Date;
					 
						Array_flightdetails_02  = Flight_Details_02.split("/");
						Carrier_code_02=Flight_Details_02.substring(0,2);
						Airline_Date_002=Array_flightdetails_02[1];
						Airline_Date_02=Airline_Date_002.substring(0,2);
						Flight_Number_02=Array_flightdetails_02[0].substring(2);
				 		Flight_date02=Airline_Date_02+Current_Date;
				}
				else if(Flight_Details_01 !="data not found")
				{
						Array_flightdetails_01  = Flight_Details_01.split("/");
						Carrier_code_01=Flight_Details_01.substring(0,2);
						Airline_Date_001=Array_flightdetails_01[1];
						Airline_Date_01=Airline_Date_001.substring(0,2);
						Flight_Number_01=Array_flightdetails_01[0].substring(2);
						Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
						Flight_date01=Airline_Date_01+Current_Date;
				}
				
//============================================================================================================================================					
				//Routing details		
				String Current_Date_for_defaultflight=new SimpleDateFormat("dd").format(Calendar.getInstance().getTime());
				driver.findElement(By.xpath(".//*[@id='imgairport']")).click();
				
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillOriginAirport_txtCode']")).sendKeys(OriPort_01);
				
				if(DestPort_01!="data not found")
				{
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']")).sendKeys(DestPort_01);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ViaRoute1_txtCode']")).sendKeys(Via1_01);
					WebElement btn_via2=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ViaRoute2_txtCode']"));
					btn_via2.sendKeys(Via2_01);
					btn_via2.sendKeys(Keys.TAB);
					Thread.sleep(1000);
					driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[8]/div[11]/button[1]")).click();
					Thread.sleep(1000);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier1']")).sendKeys(Carrier_code_01);
					if(Flight_Number_01=="data not found")
					{
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys("123");
					}
					else
					{
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys(Flight_Number_01);
					}
					
					if(Flight_date01=="data not found")
					{
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Current_Date_for_defaultflight);
					}
					else
					{
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Flight_date01);
					}
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier2']")).sendKeys(Carrier_code_02);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno2']")).sendKeys(Flight_Number_02);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate2']")).sendKeys(Flight_date02);
					
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier3']")).sendKeys(Carrier_code_03);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno3']")).sendKeys(Flight_Number_03);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate3']")).sendKeys(Flight_date03);
				}	
				else if(Via2_01!="data not found")
				{	
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']")).sendKeys(Via2_01);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ViaRoute1_txtCode']")).sendKeys(Via1_01);
					driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[8]/div[11]/button[1]")).click();
					
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier1']")).sendKeys(Carrier_code_01);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys(Flight_Number_01);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Flight_date01);
					
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier2']")).sendKeys(Carrier_code_02);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno2']")).sendKeys(Flight_Number_02);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate2']")).sendKeys(Flight_date02);
				}
				else
				{
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']")).sendKeys(Via1);
					driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[8]/div[11]/button[1]")).click();
					
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier1']")).sendKeys(Carrier_code_01);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys(Flight_Number_01);
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Flight_date01);
				}
				
//				WebElement rd=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ViaRoute2_txtName']"));
//				rd.click();
//				rd.sendKeys(Keys.TAB,Keys.ENTER);
//				Thread.sleep(1000);
			//	driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[8]/div[11]/button[1]")).click();
//==================================================================================================================================================
		//Handling Information
				String Handling_Information=eat2.getCellData("Sheet1", "Handling_Information",1);
				Handling_Information = Handling_Information.replaceAll("[^a-zA-Z0-9]", " ");
				String Handling_Information_01=Handling_Information.replaceAll("( +)"," ").trim();
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtssr']")).clear();
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtssr']")).sendKeys(Handling_Information_01);
				
				
//===================================================================================================================
				//charge code
				//	String ChargeCode=eat2.getCellData("Sheet1", "Charge_code",1);
					String ChargeCode_01="PX";
										
					WebElement cc=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ddlChargeCode']"));
					cc.click();
					Select cc1=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ddlChargeCode']")));
					cc1.selectByVisibleText(ChargeCode_01);
//============================================================================================================================================

				driver.findElement(By.xpath(".//*[@id='addDimensions_1']")).click();
				
				String Numberpieces=eat2.getCellData("Sheet1", "No_pcs",1);
				if(Numberpieces=="data not found")
				{
					driver.findElement(By.xpath(".//*[@id='txtNoPcs_1']")).sendKeys("1");
				}
				else
				{
					String Numberpieces_01 = Numberpieces.substring(0, Numberpieces.indexOf("."));
					driver.findElement(By.xpath(".//*[@id='txtNoPcs_1']")).sendKeys(Numberpieces_01);
				}
				
				String length=eat1.getCellData("AWBdetails", "Length",1);
				driver.findElement(By.xpath(".//*[@id='txtLength_1']")).sendKeys(length);
				
				String width=eat1.getCellData("AWBdetails", "Width",1);
				driver.findElement(By.xpath(".//*[@id='txtWidth_1']")).sendKeys(width);
				
				String hieght=eat1.getCellData("AWBdetails", "Height",1);
				driver.findElement(By.xpath(".//*[@id='txtHeight_1']")).sendKeys(hieght);
				
				//driver.findElement(By.xpath(".//*[@id='addrow_1']")).click();
				
				WebElement slac=driver.findElement(By.xpath(".//*[@id='txtDmnSlac_1']"));
				slac.click();
				slac.sendKeys(Keys.TAB,Keys.TAB,Keys.ENTER);
				
				
				String GrossWeight=eat2.getCellData("Sheet1", "Gross_Wt",1);
				if(GrossWeight=="data not found")
				{
					driver.findElement(By.xpath(".//*[@id='txtCgGrWt_1']")).sendKeys("1");
				}
				else
				{
					driver.findElement(By.xpath(".//*[@id='txtCgGrWt_1']")).sendKeys(GrossWeight);
				}
				
				
				String RateClass=eat2.getCellData("Sheet1", "Rate_Class",1);
				String RateClass_01=RateClass.replaceAll("( +)"," ").trim();
				if(RateClass_01.contains(""))
				{
					RateClass_01="data not found";
				}
				
				if(RateClass_01=="data not found")
				{
					Select rc1=new Select(driver.findElement(By.xpath(".//*[@id='selRateClass_1']")));
					rc1.selectByVisibleText("Q");
				}
				else
				{
					Select rc1=new Select(driver.findElement(By.xpath(".//*[@id='selRateClass_1']")));
					rc1.selectByVisibleText(RateClass_01);
				}
				
				String ComdityNo=eat2.getCellData("Sheet1", "Commodity_No",1);
				String ComdityNo_01=ComdityNo.replaceAll("( +)"," ").trim();
				if(ComdityNo_01==" ")
				{
					ComdityNo_01="data not found";
				}
				if(RateClass_01.contains("C") || RateClass_01.contains("S")) 
				{
					if(ComdityNo_01=="data not found")
					{
						driver.findElement(By.xpath(".//*[@id='txtCommNo_1']")).sendKeys("111");
					}
					else
					{
						driver.findElement(By.xpath(".//*[@id='txtCommNo_1']")).sendKeys(ComdityNo_01);
					}
				}
				
				String Charge=eat2.getCellData("Sheet1", "Charges",1);
				String Charge_01=Charge.replaceAll("( +)"," ").trim();
				
				if(Charge_01.contains("data not found") || Charge_01.contains(""))
				{
					driver.findElement(By.xpath(".//*[@id='txtCgRate_1']")).sendKeys("1");	
				}
				else
				{
				driver.findElement(By.xpath(".//*[@id='txtCgRate_1']")).sendKeys(Charge_01);
				Thread.sleep(1000);
				}
				String nature=eat2.getCellData("Sheet1", "Nature",1);
				Thread.sleep(1000);
				nature = nature.replaceAll("[^a-zA-Z0-9]", " ");
				String nature_01=nature.replaceAll("( +)"," ").trim();
				driver.findElement(By.xpath(".//*[@id='txtCgDesc_1']")).sendKeys(nature_01);
				Thread.sleep(1000);
				WebElement tcc=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSaveAwb']"));
				tcc.click();
				Thread.sleep(10000);
				
				excelsave_path=Sc;
//======================================================================================================================================================				
				endtime1 = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
		        EndTime=endtime1; 
			
		        SimpleDateFormat format = new SimpleDateFormat("HH:mm:ss");
		        Date date1 = format.parse(StartTime);
		        Date date2 = format.parse(EndTime);
		        
		        diff = date2.getTime() - date1.getTime();
		        
	            Duration=String.format("%02d", diff / hour)+":"+String.format("%02d", (diff % hour) / minute)+":"+String.format("%02d", (diff % minute) / second);
	        		
//==============================================================================================================================		
 		result3=eat1.isElementPresent(driver,".//*[@id='ctl00_hldPage_lblMessage2']");
 		Thread.sleep(3000);
		if(result3== false)//AWB not create successfully
		{  
			driver.close();
			driver.switchTo().window(tabs2.get(0));
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
    		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
    		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
    		Select status0101=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
    		status0101.selectByVisibleText("Yet To Start");
    		Thread.sleep(3000);
    		driver.switchTo().alert().accept();
    		Thread.sleep(1000);

    	      
        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtOperator']")).clear();
        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtAssignedTo']")).clear();
        	
            		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkSave']")).click();
    		Thread.sleep(5000);
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
    		Thread.sleep(2000);
    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
    		 
    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
    			Select yts0101=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
    			yts0101.selectByVisibleText("Yet To Start");        
    			i1=2;
    			
    		//	wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, excelsave_path);
    			
		}
		else//AWB create successfully
		{
			DR.read_write_D(AWBNumber, OrganizationName,DraftUpload_time,StartTime,EndTime,Duration,1, Reportsave_path);
			driver.close();
			driver.switchTo().window(tabs2.get(0));
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
    		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
    		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
    		Select status0102=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
    		status0102.selectByVisibleText("Draft Saved");
    		Thread.sleep(1000);
    		driver.switchTo().alert().accept();
    		Thread.sleep(1000);
    		
    	    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
    	    
    	    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
    	             
    		Select yts0102=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
    		yts0102.selectByVisibleText("Yet To Start");  
    		System.out.println("Airway Bill Created successfully");
    		
    		
    		long end = System.currentTimeMillis();
    		formatter = new DecimalFormat("#0.00000");
    		System.out.println(i1+".Execution time for AWB No. "+AWB_Number+"is " + formatter.format((end - start) / 1000d) + " seconds");
    		DR.foldercreate_D(Report_folder_path);
    		//wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, excelsave_path);
		}
		}
		else//shipper name not found
		{
			driver.close();
			driver.switchTo().window(tabs2.get(0));
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
    		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
    		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
    		Select status0101=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
    		status0101.selectByVisibleText("Yet To Start");
    		Thread.sleep(3000);
    		driver.switchTo().alert().accept();
    		Thread.sleep(1000);

    	      
        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtOperator']")).clear();
        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtAssignedTo']")).clear();
        	
            		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkSave']")).click();
    		Thread.sleep(5000);
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
    		Thread.sleep(2000);
    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
    		 
    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
    			Select yts0101=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
    			yts0101.selectByVisibleText("Yet To Start");        
    			i1=2;
		}
		}
    else//file not exist in target folder 
    {
    	Path temp1 = Files.move
     	        (Paths.get(Source_folder_path+renamefile1), 
     	        Paths.get(Proccess_fail_folder_path+renamefile1));
     	 
     	        if(temp1 != null)
     	        {
     	           
     	        }
     	        else
     	        {
     	            System.out.println("Failed to move the file");
     	        }
     	driver.switchTo().defaultContent();
     	Thread.sleep(3000);
    	WebElement close=driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[6]/div[11]/button"));
		Thread.sleep(3000);
		close.click();

		
		 ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());

		driver.switchTo().window(tabs2.get(1));
	   	Thread.sleep(3000);
	    driver.close();
	    
	   	driver.switchTo().window(tabs2.get(0));
	   
	   	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
		
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
		
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
		Select status0104=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
		status0104.selectByVisibleText("Yet To Start");
		Thread.sleep(2000);
		driver.switchTo().alert().accept();
		Thread.sleep(1000);

	      
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtOperator']")).clear();
    	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtAssignedTo']")).clear();
    	
        		
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkSave']")).click();
		Thread.sleep(3000);
		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
		Thread.sleep(2000);
	    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
	    
	    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
		Select yts0104=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
		yts0104.selectByVisibleText("Yet To Start");        
    }
     	}  
     
                		}
					                else
					                	{
					            			i1++;
					                	}
            			}
						            else
						        		{
						            		i1++;
						        		}
						}
							        else
							        	{
							        		i1++;
							        	}
        	 			}
						        	else
							        	{
							        		i1++;
							        	}
        	 			}
						        	else
						        		{
						        		i1++;
						        		}
//        	 			}
//						        	else
//						        		{
//						        		i1++;
//						        		}
        					
    
        }
        
       
//        	Reason="There is no such Drafts present";
//        	wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, filename);
        }
			Reason="There is no such Drafts present";
			
//			driver.close();
//			wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, filename);
	}
		
//***************************************************************//*********************************************//****************************
//****************************************************************//**********************************************//**************************		
		else
		{
			List<WebElement> all = driver.findElements(By.xpath("//input[contains(@id,'txtOtherComments')]"));
	        String[] allText = new String[all.size()];
	        int i1 = 2; 

	        for (WebElement element : all)
	        {	
//	        	result4=eat1.isElementPresent(driver, ".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgAssgnDateTime']");
//	        	if(result4==false) 
//	        	{
	        	result=eat1.isElementPresent(driver, ".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0" + i1 + "_imgFinalUpload']");
	        	if(result==false)
	        	{
	        	WebElement housecount=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtHAWBCount']"));
	        	 String count = housecount.getAttribute("value");
	        
	        	 if(count.contains("0"))
	        	 {
	        		 result1=eat1.isElementPresent(driver, ".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0" + i1 + "_imgDraftUpload']");
	       
	            if (result1== true)
				{
	            	result2=eat1.isElementPresent(driver,".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgAWBCreation']");
	            	
	   
	            
	            	if(result2==false)
	            {
	            		Actions draftup = new Actions(driver);
	                	WebElement imgDraftUpload = driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_imgDraftUpload']"));
	                	Thread.sleep(2000);
	                	draftup.clickAndHold(imgDraftUpload).perform();
	                	String DraftUploadDate_Time = imgDraftUpload.getAttribute("title");
	                	
	                	String[] arrOfStr1 = DraftUploadDate_Time.split(" ");
	    				String DraftUpload_time = arrOfStr1[1];
	                	            		
	            		starttime1 = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
	      		        StartTime=starttime1;
	      		        
	            		WebElement AWBno=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_lblawbno']"));
	            		String AWBNumber=AWBno.getText();
	            	
	            		String [] WholeAWB  = AWBNumber.split("-");
	            		String onlyPrefix = WholeAWB [0];
	            		String onlyAWB=WholeAWB [1];
	            		AWB_Number=onlyAWB;
	            		
	            		onlyPrefix = onlyPrefix.replaceAll(" ", "");
	            	
	            		WebElement OrgName=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack']/tbody/tr["+i1+"]/td[17]"));
	            		String OrganizationName=OrgName.getText();
	            		System.out.println(OrganizationName);

	            	// excel_operations eat3=new excel_operations("D:\\VIVEK\\TIFFA\\TestData\\RuleExcel.xlsx"); 
	                 int r=0; String RuleName=null;
	                for(r=1;r<=59;r++)
	                {
	             	 String OrganiseName=eat3.getCellData("RuleMaster",r,"Agent_Name");	
	             	 
	             	 if(OrganiseName.equals(OrganizationName))
	             	 {
	             		RuleName=eat3.getCellData("RuleMaster",r,"Rule_Name");	
	             	 }
	                }
	                if(!RuleName.equals("Scan") && !RuleName.equals("Not_Done") )
	                {
	                	String UpdateRuleName=RuleName+"_"+onlyPrefix;
	                	//String UpdateRuleName="SIAM_098";
	                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_drpLinks']")).click();
	            		Select status=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_drpLinks']")));
	            		status.selectByVisibleText("Assigned");
	            		Thread.sleep(5000);
	            		driver.switchTo().alert().accept();
	            		Thread.sleep(5000);
	                	
	            		if(username.contains("tiffaawb"))
	            	    {
	            	    	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtOperator']")).sendKeys("testair");
	                    	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtAssignedTo']")).sendKeys("testair");
	            	    }
	            	    else
	            	    {
	                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtOperator']")).sendKeys("Automation");
	                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_txtAssignedTo']")).sendKeys("Automation");
	            	    }
	                	
	            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl0"+i1+"_lnkSave']")).click();
	            		Thread.sleep(10000);
	            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
	            		Thread.sleep(10000);
	            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
	            		
	            		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
	            		
	            		
	            		//Thread.sleep(2000);

	                	
	                	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkViewLatest']")).click();
	                    Thread.sleep(10000);
	            		driver.switchTo().frame("Ifram1");
	            		Thread.sleep(10000);
	            		
	            		driver.findElement(By.linkText("View")).click();
	            		Thread.sleep(10000);
	                     
	//=============================================================================================================================================
	            		/**
	          		   * Files extraction.
	          		   */
	                  String Chrome_Downloads_path=props.getProperty("Chrome_Downloads_path");  
	                  String Source_folder_path=props.getProperty("Source_folder_path"); 
	                  String Target_folder_path=props.getProperty("Target_folder_path"); 
	                  String Proccessed_folder_path=props.getProperty("Proccessed_folder_path"); 
	                  String Proccess_fail_folder_path=props.getProperty("Proccess_fail_folder_path");
	          		final File folder = new File(Chrome_Downloads_path);
	               String fileName = "null";
	               for (final File fileEntry : folder.listFiles())
	              {
	                   	  listFilesForFolder(fileEntry);
	                       fileName=fileEntry.getName();
	//===============================================================================================================================================
	             Path temp = Files.move
	            		 (Paths.get(Chrome_Downloads_path+fileName), 
	                    	        Paths.get(Source_folder_path+fileName));
	           	 
	           	        if(temp != null)
	           	        {
	           	            
	           	        }
	           	        else
	           	        {
	           	            System.out.println("Failed to move the file");
	           	        }
	             
	//===============================================================================================================================================              
	           	     File file = new File(Source_folder_path+fileName);
	                 String renamefile=new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
	                 String renamefile1=renamefile+".pdf";
	                 File input_file = new File(Source_folder_path+renamefile1);
	                 if(file.renameTo(input_file))
	             {
	                 
	             }
	             else
	             {
	                 System.out.println("File rename failed");
	             }
	//===============================================================================================================================================              
	               
	                 String targetfile=new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
	                 String targetfile1=targetfile+".xls";
	                 String TARGET_FILE=Target_folder_path+targetfile1;
	//===========================================================================================================================
	                 try
	                 { 
	                 
	               	//  Runtime.getRuntime().exec("cmd /c start PDECMD -R\"OCR_PDF_RULE\" -F\""+newFile+"\" -O\"D:\\VIVEK\\TIFFA\\OCR\\Target\\pdfdata -TXLS -PA.xls");
	               	 // Runtime.getRuntime().exec("cmd /c start PDECMD -R\"SIAM_KARGO_LOGISTICS\" -F\""+newFile+"\" -O\"D:\\VIVEK\\TIFFA\\OCR\\Target\\"+targetfile1);
	                	 Runtime.getRuntime().exec("cmd /c start PDECMD -R\""+UpdateRuleName+"\" -F\""+input_file+"\" -O\""+TARGET_FILE);
	               	  Thread.sleep(8000);
	                 }
	                 
	                 catch (Exception e)
	                 {
	                     System.out.println("Incorrect action perform");
	                     e.printStackTrace();
	                 }
	//========================================================================================================================================================	
	                 File f = new File(TARGET_FILE);
	                 if(f.exists() && !f.isDirectory()) 
	                 { 
	                     
	                 
	//========================================================================================================================================================	
	     Path temp1 = Files.move
	    		 (Paths.get(Source_folder_path+renamefile1), 
	    		 	        Paths.get(Proccessed_folder_path+renamefile1));	 	
	     
	 	        if(temp1 != null)
	 	        {
	 	           
	 	        }
	 	        else
	 	        {
	 	            System.out.println("Failed to move the file");
	 	        }
	//=========================================================================================================================================================      
	 	       excel_operations eat2=new excel_operations(Target_folder_path+targetfile1);
				driver.switchTo().defaultContent();
				Thread.sleep(5000);
				WebElement close=driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[6]/div[11]/button"));
				Thread.sleep(10000);
				close.click();

				
				 ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());
				 
//				 driver.close();
//				 Thread.sleep(3000);
			
			   driver.switchTo().window(tabs2.get(1));
			
				    Thread.sleep(10000);
			  
			long start = System.currentTimeMillis();
			
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBPrefix']")).sendKeys(onlyPrefix);			
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(String.valueOf(onlyAWB));
			String ShipName=eat2.getCellData("Sheet1", "Shipper_Name",1);
			if(!ShipName.equals("data not found"))
			{
//========================================================================================================================================
			driver.findElement(By.xpath(".//*[@id='btnshipperadd']")).click();
			
			
			ShipName = ShipName.replaceAll("[^a-zA-Z0-9]", " ");
			String ShipName_01=ShipName.replaceAll("( +)"," ").trim();
			
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCompanyName']")).sendKeys(ShipName_01);
			
//			int ShipName_length=ShipName_01.length();
//			if(ShipName_length >35 )
//			{
//			String limit_ShipName = ShipName_01.substring(0, 35);
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCompanyName']")).sendKeys(limit_ShipName);
//			}
//			else
//			{
//				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCompanyName']")).sendKeys(ShipName);
//			}
//============================================================================================================================================
			String Remain_Addline="";String Concate_add2_add1=null;
			String Addline_01=eat2.getCellData("Sheet1", "Address_Line01",1);
			Thread.sleep(1000);
			Addline_01 = Addline_01.replaceAll("[^a-zA-Z0-9]", " ");
			String MAddline_01=Addline_01.replaceAll("( +)"," ").trim();
			
			String Addline_02=eat2.getCellData("Sheet1", "Address_Line02",1);
			String Concate_AddLine02;
			Thread.sleep(1000);
			Addline_02 = Addline_02.replaceAll("[^a-zA-Z0-9]", " ");
			String MAddline_02=Addline_02.replaceAll("( +)"," ").trim();
			
			
			if(MAddline_01=="data not found" || MAddline_01=="" || MAddline_01==" ")
			{
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 1");	
			}
			else
			{
			int Addline_length =MAddline_01.length();
			
			if(Addline_length >35)
			{
					String limit_Addline = MAddline_01.substring(0, 35);
					 Remain_Addline=MAddline_01.substring(36, Addline_length);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys(limit_Addline);
			}
			else
			{
				Thread.sleep(1000);
				
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys(MAddline_01);
			}
			}
//=======================================================================================================================================
			
			if(MAddline_02=="data not found" || MAddline_02==" " || MAddline_02=="")
			{
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 2");
			}
			else
			{
				if(Remain_Addline=="")
				{
					Concate_AddLine02=MAddline_02;
				}
				else
				{
					Concate_AddLine02=Remain_Addline+MAddline_02;
				}
			
			int Addline_length_02 =Concate_AddLine02.length();
			if(Addline_length_02 >35)
			{
					
			String limit_Addline_02=Concate_AddLine02.substring(0,35);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine2']")).sendKeys(limit_Addline_02);
			}
			else
			{
				Thread.sleep(1000);
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine2']")).sendKeys(Concate_AddLine02);
			}
			}
//===============================================================================================================================================			
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCountry_txtCode']")).sendKeys("IN");
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillState_txtCode']")).sendKeys("MH");
//
//			Thread.sleep(1000);
////			String CityName=eat2.getCellData("Sheet1", "City",2);
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCity_txtCode']")).sendKeys("BOM");
//			Thread.sleep(1000);
//
////			String PinCodeNo=eat2.getCellData("Sheet1", "PinCode",1);
//			WebElement we=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtPinCode']"));
//			we.sendKeys("400610");
//			we.click();
//			
//
//	       
//			Thread.sleep(1000);
//			we.sendKeys(Keys.TAB,Keys.ENTER);
//=====================================================================================================================================================
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCountry_txtCode']")).sendKeys("TH");
			

			Thread.sleep(1000);
//			String CityName=eat2.getCellData("Sheet1", "City",2);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_GenericAutoFillCity_txtCode']")).sendKeys("BKK");
			Thread.sleep(1000);

//			String PinCodeNo=eat2.getCellData("Sheet1", "PinCode",1);
			WebElement we=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtPinCode']"));
			we.sendKeys("XXXXXX");
			we.click();
			

	       
			Thread.sleep(1000);
			we.sendKeys(Keys.TAB,Keys.ENTER);
	//===============================================================================================================================================		
			
			driver.findElement(By.xpath(".//*[@id='imgbtnConsignee']")).click();
			
			String ConsigName=eat2.getCellData("Sheet1", "Consignee_Name",1);
			
			
			ConsigName = ConsigName.replaceAll("[^a-zA-Z0-9]", " ");
			String ConsigName_01=ConsigName.replaceAll("( +)"," ").trim();
			int ConsigName_length=ConsigName_01.length();
			if(ConsigName_length >35)
			{
			String limit_ConsigName = ConsigName_01.substring(0, 35);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtConName']")).sendKeys(limit_ConsigName);
			}
			else
			{
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtConName']")).sendKeys(ConsigName);	
			}
//=================================================================================================================================================			
			String Remain_Addline_02="";
			String CAddline_01=eat2.getCellData("Sheet1", "Address_Line01_c",1);
			
			Thread.sleep(1000);
			CAddline_01 = CAddline_01.replaceAll("[^a-zA-Z0-9]", " ");
			String CMAddline_01=CAddline_01.replaceAll("( +)"," ").trim();
			if(CMAddline_01=="data not found" || CMAddline_01=="" || CMAddline_01==" ")
			{
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 1");	
			}
			else
			{
			int CAddline1_length=CMAddline_01.length();
			if(CAddline1_length >35)
			{
			String limit_Addline01 = CMAddline_01.substring(0, 35);
			Remain_Addline_02=CMAddline_01.substring(36, CAddline1_length);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine1']")).sendKeys(limit_Addline01);
			}
			else
			{
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine1']")).sendKeys(CMAddline_01);
			}
			}
//====================================================================================================================================================
			String CAddline_02=eat2.getCellData("Sheet1", "Address_Line02_c",1);
			String Concate_CAddLine02;
			Thread.sleep(1000);
			CAddline_02 = CAddline_02.replaceAll("[^a-zA-Z0-9]", " ");
			String CMAddline_02=CAddline_02.replaceAll("( +)"," ").trim();
			if(CMAddline_02=="data not found" || CMAddline_02=="" || CMAddline_02==" ")
			{
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclOrgAddressInformation_txtAddressLine1']")).sendKeys("DUMMY Address line 2");	
			}
			else
			{
				if(Remain_Addline_02=="")
				{
					Concate_CAddLine02=CMAddline_02;
				}
				else
				{
					Concate_CAddLine02=Remain_Addline_02+CMAddline_02;
				}
				
			int CAddline2_length=Concate_CAddLine02.length();
			if(CAddline2_length >35)
			{
			String limit_Addline02 = Concate_CAddLine02.substring(0, 35);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine2']")).sendKeys(limit_Addline02);
			}
			else
			{
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtAddressLine2']")).sendKeys(Concate_CAddLine02);
			}
			}			
//===============================================================================================================================================
			//String CountryName1=eat1.getCellData("AWBdetails", "Country_c",1);
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCountry_txtName']")).sendKeys("United Arab Emirates");
//			
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillState_txtCode']")).sendKeys("DX");
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCity_txtCode']")).sendKeys("DXB");
//			
//			//String CityName1=eat2.getCellData("Sheet1", "City_c",2);
//			//driver.findElement(By.id("ctl00_hldPage_uclConsAddressInfo_txtOtherCity")).sendKeys("MUMBAI");
//			
//			
//			//String PinCode1=eat2.getCellData("Sheet1", "PinCode_c",2);
//			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtPinCode']")).sendKeys("000000");
//			Thread.sleep(1000);
//			
//			WebElement wf=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclAlsoNotifyAddressInfo_txtPinCode']"));
//			wf.click();
//			wf.sendKeys(Keys.TAB,Keys.ENTER);
//			Thread.sleep(1000);
//==============================================================================================================================================
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCountry_txtName']")).sendKeys("INDIA");
			
			
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_GenericAutoFillCity_txtCode']")).sendKeys("BOM");
			
			//String CityName1=eat2.getCellData("Sheet1", "City_c",2);
			//driver.findElement(By.id("ctl00_hldPage_uclConsAddressInfo_txtOtherCity")).sendKeys("MUMBAI");
			
			
			//String PinCode1=eat2.getCellData("Sheet1", "PinCode_c",2);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclConsAddressInfo_txtPinCode']")).sendKeys("XXXXXX");
			Thread.sleep(1000);
			
			WebElement wf=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_uclAlsoNotifyAddressInfo_txtPinCode']"));
			wf.click();
			wf.sendKeys(Keys.TAB,Keys.ENTER);
			Thread.sleep(1000);
			
			
	//===============================================================================================================================================
	//Issuing Carrier's Agent Name and City
			String IssuAgent=eat2.getCellData("Sheet1", "Agent",1);
			
			WebElement agnt=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAgentSelect']"));
//			if(IssuAgent.equals("data not found") && IssuAgent.equals("  "))
//			{
			if(OrganizationName.equalsIgnoreCase("TIFFA AWB Service"))
			{
				IssuAgent="KALE LOGISTICS";
			}
			else
			{
			IssuAgent=OrganizationName;
			}
				//IssuAgent="EAGLES AIR AND SEA (THAILAND) CO., LTD";


				
		//	}
			//String AgentConcate= IssuAgent.substring(0, IssuAgent.indexOf("."));
			
			agnt.sendKeys(IssuAgent);
			Thread.sleep(2000);
			agnt.sendKeys(Keys.CONTROL+"A");
			Thread.sleep(2000);
			agnt.sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(1000);
			agnt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			
			String AccountInformatiom=eat2.getCellData("Sheet1", "Accounting_Information",1);
			AccountInformatiom = AccountInformatiom.replaceAll("[^a-zA-Z0-9]", " ");
			String AccountInformatiom_01=AccountInformatiom.replaceAll("( +)"," ").trim();
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_accountinginfo']")).sendKeys(AccountInformatiom_01);
			//================================================================================================================================
			//flight details
			String Carrier_code_01=null,Airline_Date_01,Airline_Date_001,Flight_Number_01=null,Current_Date,Flight_date01=null,Carrier_code_02=null,Airline_Date_02=null,Airline_Date_002,Flight_Number_02=null,Flight_date02=null,Carrier_code_03=null,Airline_Date_03=null,Airline_Date_003,Flight_Number_03=null,Flight_date03=null;
			 String [] Array_flightdetails_01,Array_flightdetails_02,Array_flightdetails_03;
					 Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
					 
					 String OriPort=eat2.getCellData("Sheet1", "Origin_Port",1);
					 String OriPort_01=OriPort.replaceAll("( +)"," ").trim();
					 
					 if(OriPort_01=="" ||  OriPort=="data not found")
					 {
						 OriPort_01="data not found"; 
					 }
					 
					 if(OriPort_01.equals("SUVARNABHUMI AIRPORT , THAILAND"))
					 {
						 OriPort_01="BKK";
					 }
					 else if(OriPort_01.equals("BANGKOK/THAILAND"))
					 {
						 OriPort_01="BKK";
					 }
					 else if(OriPort_01.equals("BANGKOK,THAILAND"))
					 {
						 OriPort_01="BKK";
					 }
					 else if(OriPort_01.equals("SUVARNABHUMI AIRPORT"))
					 {
						 OriPort_01="BKK";
					 }
					 else if(OriPort_01.equals("DON MUEANG AIRPORT, THAILAND"))
					 {
						 OriPort_01="DMK";
					 }
					 else if(OriPort_01.equals("DON MUEANG AIRPORT"))
					 {
						 OriPort_01="DMK";
					 }
					 
					String Via1=eat2.getCellData("Sheet1", "Via01",1);
					String Via1_01=Via1.replaceAll("( +)"," ").trim();
					if(Via1_01=="" || Via1=="data not found")
					 {
						Via1_01="data not found"; 
					 }
					
					String Via2=eat2.getCellData("Sheet1", "Via02",1);
					String Via2_01=Via2.replaceAll("( +)"," ").trim();
					
					if(Via2_01=="" || Via2=="data not found")
					 {
						Via2_01="data not found"; 
					 }
					
					String DestPort=eat2.getCellData("Sheet1", "Destn_Port",1);
					String DestPort_01=DestPort.replaceAll("( +)"," ").trim();
					if(DestPort_01.contentEquals("") || DestPort.contentEquals("data not found"))
					 {
						DestPort_01="data not found"; 
					 }
					
			
					String Flight_Details_01=eat2.getCellData("Sheet1", "Flight_Details_01",1);
					String Flight_Details_02=eat2.getCellData("Sheet1", "Flight_Details_02",1);
					String Flight_Details_03=eat2.getCellData("Sheet1", "Flight_Details_03",1);
					

					if(IssuAgent=="KALE LOGISTICS")
					{
						Flight_Details_01="9W";
						Flight_Details_02="9W1922/26";
						Flight_Details_03="data not found";
					}
					int lenght_fli_det_01=Flight_Details_01.length();
					if(lenght_fli_det_01<=4)
					{
						Flight_Details_01="data not found";
					}
					
					
					if(DestPort_01=="data not found")
					{
						Flight_Details_01=Flight_Details_02;
						Flight_Details_02=Flight_Details_03;
					}
					else if(Via2=="data not found")
					{
						Flight_Details_01=Flight_Details_02;
					}
					
					if(Flight_Details_03 !="data not found")
					{				
						Array_flightdetails_01  = Flight_Details_01.split("/");
						 Carrier_code_01=Flight_Details_01.substring(0,2);
						 Airline_Date_001=Array_flightdetails_01[1];
						 Airline_Date_01=Airline_Date_001.substring(0,2);
						 Flight_Number_01=Array_flightdetails_01[0].substring(2);
						 Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
						 Flight_date01=Airline_Date_01+Current_Date;
					 
						 Array_flightdetails_02  = Flight_Details_02.split("/");
						 Carrier_code_02=Flight_Details_02.substring(0,2);
						 Airline_Date_002=Array_flightdetails_02[1];
						 Airline_Date_02=Airline_Date_002.substring(0,2);
						 Flight_Number_02=Array_flightdetails_02[0].substring(2);
						 Flight_date02=Airline_Date_02+Current_Date;
					
						 Array_flightdetails_03  = Flight_Details_03.split("/");
						 Carrier_code_03=Flight_Details_03.substring(0,2);
						 Airline_Date_003=Array_flightdetails_03[1];
						 Airline_Date_03=Airline_Date_003.substring(0,2);
						 Flight_Number_03=Array_flightdetails_03[0].substring(2);
						 Flight_date03=Airline_Date_03+Current_Date;
					}

					else if(Flight_Details_02 !="data not found")
					{
							Array_flightdetails_01  = Flight_Details_01.split("/");
							Carrier_code_01=Flight_Details_01.substring(0,2);
							Airline_Date_001=Array_flightdetails_01[1];
							Airline_Date_01=Airline_Date_001.substring(0,2);
							Flight_Number_01=Array_flightdetails_01[0].substring(2);
							Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
							Flight_date01=Airline_Date_01+Current_Date;
						 
							Array_flightdetails_02  = Flight_Details_02.split("/");
							Carrier_code_02=Flight_Details_02.substring(0,2);
							Airline_Date_002=Array_flightdetails_02[1];
							Airline_Date_02=Airline_Date_002.substring(0,2);
							Flight_Number_02=Array_flightdetails_02[0].substring(2);
					 		Flight_date02=Airline_Date_02+Current_Date;
					}
					else if(Flight_Details_01 !="data not found")
					{
							Array_flightdetails_01  = Flight_Details_01.split("/");
							Carrier_code_01=Flight_Details_01.substring(0,2);
							Airline_Date_001=Array_flightdetails_01[1];
							Airline_Date_01=Airline_Date_001.substring(0,2);
							Flight_Number_01=Array_flightdetails_01[0].substring(2);
							Current_Date=new SimpleDateFormat("/MM/yyyy").format(Calendar.getInstance().getTime());
							Flight_date01=Airline_Date_01+Current_Date;
					}
					
	//================================================================================================================================
	//Routing details		
					String Current_Date_for_defaultflight=new SimpleDateFormat("dd").format(Calendar.getInstance().getTime());
					driver.findElement(By.xpath(".//*[@id='imgairport']")).click();
					
					driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillOriginAirport_txtCode']")).sendKeys(OriPort_01);
					
					if(DestPort_01!="data not found")
					{
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']")).sendKeys(DestPort_01);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ViaRoute1_txtCode']")).sendKeys(Via1_01);
						WebElement btn_via2=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ViaRoute2_txtCode']"));
						btn_via2.sendKeys(Via2_01);
						btn_via2.sendKeys(Keys.TAB);
						Thread.sleep(1000);
						driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[8]/div[11]/button[1]")).click();
						Thread.sleep(1000);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier1']")).sendKeys(Carrier_code_01);
						if(Flight_Number_01=="data not found")
						{
							driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys("123");
						}
						else
						{
							driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys(Flight_Number_01);
						}
						
						if(Flight_date01=="data not found")
						{
							driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Current_Date_for_defaultflight);
						}
						else
						{
							driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Flight_date01);
						}
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier2']")).sendKeys(Carrier_code_02);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno2']")).sendKeys(Flight_Number_02);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate2']")).sendKeys(Flight_date02);
						
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier3']")).sendKeys(Carrier_code_03);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno3']")).sendKeys(Flight_Number_03);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate3']")).sendKeys(Flight_date03);
					}	
					else if(Via2_01!="data not found")
					{	
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']")).sendKeys(Via2_01);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ViaRoute1_txtCode']")).sendKeys(Via1_01);
						driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[8]/div[11]/button[1]")).click();
						
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier1']")).sendKeys(Carrier_code_01);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys(Flight_Number_01);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Flight_date01);
						
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier2']")).sendKeys(Carrier_code_02);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno2']")).sendKeys(Flight_Number_02);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate2']")).sendKeys(Flight_date02);
					}
					else
					{
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_GenericAutoFillDestAirport_txtCode']")).sendKeys(Via1);
						driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[8]/div[11]/button[1]")).click();
						
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtCarrier1']")).sendKeys(Carrier_code_01);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightno1']")).sendKeys(Flight_Number_01);
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtflightdate1']")).sendKeys(Flight_date01);
					}
//==================================================================================================================================================
	//Handling Information
					String Handling_Information=eat2.getCellData("Sheet1", "Handling_Information",1);
						Handling_Information = Handling_Information.replaceAll("[^a-zA-Z0-9]", " ");
						String Handling_Information_01=Handling_Information.replaceAll("( +)"," ").trim();
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtssr']")).clear();
						driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtssr']")).sendKeys(Handling_Information_01);
												
		
			
	//===================================================================================================================
					//charge code
				//	String ChargeCode=eat2.getCellData("Sheet1", "Charge_code",1);
					String ChargeCode_01="PX";
										
					WebElement cc=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ddlChargeCode']"));
					cc.click();
					Select cc1=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_ddlChargeCode']")));
					cc1.selectByVisibleText(ChargeCode_01);
	//============================================================================================================================================
					
					driver.findElement(By.xpath(".//*[@id='addDimensions_1']")).click();
					
					String Numberpieces=eat2.getCellData("Sheet1", "No_pcs",1);
					if(Numberpieces=="data not found")
					{
						driver.findElement(By.xpath(".//*[@id='txtNoPcs_1']")).sendKeys("1");
					}
					else
					{
						String Numberpieces_01 = Numberpieces.substring(0, Numberpieces.indexOf("."));
						driver.findElement(By.xpath(".//*[@id='txtNoPcs_1']")).sendKeys(Numberpieces_01);
					}
					
					String length=eat1.getCellData("AWBdetails", "Length",1);
					driver.findElement(By.xpath(".//*[@id='txtLength_1']")).sendKeys(length);
					
					String width=eat1.getCellData("AWBdetails", "Width",1);
					driver.findElement(By.xpath(".//*[@id='txtWidth_1']")).sendKeys(width);
					
					String hieght=eat1.getCellData("AWBdetails", "Height",1);
					driver.findElement(By.xpath(".//*[@id='txtHeight_1']")).sendKeys(hieght);
					
					//driver.findElement(By.xpath(".//*[@id='addrow_1']")).click();
					
					WebElement slac=driver.findElement(By.xpath(".//*[@id='txtDmnSlac_1']"));
					slac.click();
					slac.sendKeys(Keys.TAB,Keys.TAB,Keys.ENTER);
			
			
			String GrossWeight=eat2.getCellData("Sheet1", "Gross_Wt",1);
			if(GrossWeight=="data not found")
			{
				driver.findElement(By.xpath(".//*[@id='txtCgGrWt_1']")).sendKeys("1");
			}
			else
			{
				driver.findElement(By.xpath(".//*[@id='txtCgGrWt_1']")).sendKeys(GrossWeight);
			}
			
			
              			String RateClass=eat2.getCellData("Sheet1", "Rate_Class",1);
			String RateClass_01=RateClass.replaceAll("( +)"," ").trim();
			if(RateClass_01.contains(""))
			{
				RateClass_01="data not found";
			}
			
			if(RateClass_01=="data not found")
			{
				Select rc1=new Select(driver.findElement(By.xpath(".//*[@id='selRateClass_1']")));
				rc1.selectByVisibleText("Q");
			}
			else
			{
				Select rc1=new Select(driver.findElement(By.xpath(".//*[@id='selRateClass_1']")));
				rc1.selectByVisibleText(RateClass_01);
			}
			
			String ComdityNo=eat2.getCellData("Sheet1", "Commodity_No",1);
			String ComdityNo_01=ComdityNo.replaceAll("( +)"," ").trim();
			if(ComdityNo_01==" ")
			{
				ComdityNo_01="data not found";
			}
			if(RateClass_01.contains("C") || RateClass_01.contains("S")) 
			{
				if(ComdityNo_01=="data not found")
				{
					driver.findElement(By.xpath(".//*[@id='txtCommNo_1']")).sendKeys("111");
				}
				else
				{
					driver.findElement(By.xpath(".//*[@id='txtCommNo_1']")).sendKeys(ComdityNo_01);
				}
			}
			
			String Charge=eat2.getCellData("Sheet1", "Charges",1);
			String Charge_01=Charge.replaceAll("( +)"," ").trim();
			if(Charge_01.contentEquals("data not found") || Charge_01.contentEquals(""))
			{
				driver.findElement(By.xpath(".//*[@id='txtCgRate_1']")).sendKeys("1");	
			}
			else
			{
			driver.findElement(By.xpath(".//*[@id='txtCgRate_1']")).sendKeys(Charge_01);
			Thread.sleep(1000);
			}
             String nature=eat2.getCellData("Sheet1", "Nature",1);
			Thread.sleep(1000);
			nature = nature.replaceAll("[^a-zA-Z0-9]", " ");
			String nature_01=nature.replaceAll("( +)"," ").trim();
			driver.findElement(By.xpath(".//*[@id='txtCgDesc_1']")).sendKeys(nature_01);
			Thread.sleep(1000);
			WebElement tcc=driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSaveAwb']"));
			tcc.click();
			Thread.sleep(10000);
			
			Sc=eat1.screenShot(driver);
			excelsave_path=Sc;
//==============================================================================================================================
			endtime1 = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
	        EndTime=endtime1; 
		
	        SimpleDateFormat format = new SimpleDateFormat("HH:mm:ss");
	        Date date1 = format.parse(StartTime);
	        Date date2 = format.parse(EndTime);
	        
	        diff = date2.getTime() - date1.getTime();
	        
            Duration=String.format("%02d", diff / hour)+":"+String.format("%02d", (diff % hour) / minute)+":"+String.format("%02d", (diff % minute) / second);
        						
//==============================================================================================================================		
			result3=eat1.isElementPresent(driver,".//*[@id='ctl00_hldPage_lblMessage2']");
			Thread.sleep(3000);
			if(result3== false)
			{
				driver.close();
				driver.switchTo().window(tabs2.get(0));
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
	    		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
	    		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
	    		Select status0201=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
	    		status0201.selectByVisibleText("Yet To Start");
	    		Thread.sleep(3000);
	    		driver.switchTo().alert().accept();
	    		Thread.sleep(1000);

	    	      
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtOperator']")).clear();
	        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtAssignedTo']")).clear();
	        	
	            		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkSave']")).click();
	    		Thread.sleep(5000);
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
	    		Thread.sleep(2000);
	    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
	    		 
	    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
	    			Select yts0201=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
	    			yts0201.selectByVisibleText("Yet To Start"); 
	    			i1=2;
	    		//	wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, excelsave_path);
			}
			else
			{
				DR.read_write_D(AWBNumber, OrganizationName,DraftUpload_time,StartTime,EndTime,Duration,1, Reportsave_path);
				driver.close();
				driver.switchTo().window(tabs2.get(0));
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
	    		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
	    		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
	    		Select status0102=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
	    		status0102.selectByVisibleText("Draft Saved");
	    		Thread.sleep(1000);
	    		driver.switchTo().alert().accept();
	    		Thread.sleep(1000);
	    		
	    	    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
	    	    
	    	    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
	    	             
	    		Select yts0202=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
	    		yts0202.selectByVisibleText("Yet To Start");  
	    		System.out.println("Airway Bill Created successfully");
	    		long end = System.currentTimeMillis();
	    		//NumberFormat formatter = new DecimalFormat("#0.00");
	    		System.out.println(i1+".Execution time for AWB No. "+AWB_Number+"is " + formatter.format((end - start) / 1000d) + " seconds");  
	    		
	    		//wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, excelsave_path);
			}
			}
			else//shipper name not found
			{
				driver.close();
				driver.switchTo().window(tabs2.get(0));
				driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
	    		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
	    		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
	    		Select status0101=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
	    		status0101.selectByVisibleText("Yet To Start");
	    		Thread.sleep(3000);
	    		driver.switchTo().alert().accept();
	    		Thread.sleep(1000);

	    	      
	        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtOperator']")).clear();
	        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtAssignedTo']")).clear();
	        	
	            		
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkSave']")).click();
	    		Thread.sleep(5000);
	    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
	    		Thread.sleep(2000);
	    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
	    		 
	    		 driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
	    			Select yts0101=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
	    			yts0101.selectByVisibleText("Yet To Start");        
	    			i1=2;
			}

			}
			
	    
	    else
	    {
	    	Path temp1 = Files.move
	     	        (Paths.get(Source_folder_path+renamefile1), 
	     	        Paths.get(Proccess_fail_folder_path+renamefile1));
	     	 
	     	        if(temp1 != null)
	     	        {
	     	           
	     	        }
	     	        else
	     	        {
	     	            System.out.println("Failed to move the file");
	     	        }
	     	driver.switchTo().defaultContent();
	     	Thread.sleep(3000);
	    	WebElement close=driver.findElement(By.xpath(".//*[@id='aspnetForm']/div[6]/div[11]/button"));
			Thread.sleep(3000);
			close.click();

			
			 ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());

			driver.switchTo().window(tabs2.get(1));
		   	Thread.sleep(3000);
		   	driver.findElement(By.xpath(".//*[@id='ctl00_lnkSignout']")).click();
		    driver.close();
		    
		   	driver.switchTo().window(tabs2.get(0));
		   
		   	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_txtAWBNo']")).sendKeys(onlyAWB);
			
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnSearchAWB']")).click();
			
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")).click();
			Select status0204=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_drpLinks']")));
			status0204.selectByVisibleText("Yet To Start");
			Thread.sleep(1000);
			driver.switchTo().alert().accept();
			Thread.sleep(1000);

		      
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtOperator']")).clear();
        	driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_txtAssignedTo']")).clear();
        	
            		
    		driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_grdAWBTrack_ctl02_lnkSave']")).click();
			Thread.sleep(3000);
			driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnCloseDialog']")).click();
			Thread.sleep(2000);
		    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_btnLoadAll']")).click();
		    
		    driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")).click();              // Click Dropdown 
			Select yts0204=new Select(driver.findElement(By.xpath(".//*[@id='ctl00_hldPage_drpStatus']")));
			yts0204.selectByVisibleText("Yet To Start");        
	    }
	     	}  
	     
	                		}
						                else
						                	{
						            			i1++;
						                	}
	            			}
							            else
							        		{
							            		i1++;
							        		}
							}
								        else
								        	{
								        		i1++;
								        	}
	        	 			}
							        	else
								        	{
								        		i1++;
								        	}
	        	 			}
							        	else
							        		{
							        			i1++;
							        		}
//	        	 			}
//							        	else
//							        		{
//							        			i1++;
//							        		}
	        					
	    
	        }
	        }
				System.out.println("There is no such Drafts present");
				Thread.sleep(5000);
				driver.findElement(By.xpath(".//*[@id='ctl00_lnkSignout']")).click();
				Thread.sleep(5000);
				driver.close();
				driver.quit();
			//	wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, excelsave_path);
		}
		else
		{
		System.out.println("Execution is successfully completed");
		Thread.sleep(3000);
		driver.findElement(By.xpath(".//*[@id='ctl00_lnkSignout']")).click();
		Thread.sleep(5000);
		driver.close();
		driver.quit();
		}
		}
		catch(Exception e)
		{
		System.out.println(e.getMessage());
		String reason=e.getMessage();
		Mail mai=new Mail();
		Status="Fail";
		
		String ExcelPath="E:\\Tiffa_Project\\TestExecution\\Test_ExecutionREPORT.xls";
       mai.SendMail(AWB_Number, Status,ExcelPath,reason);
       wx.read_write(Srno, AWB_Number, Agent_Name, starttime1, endtime1, Duration, Status, Reason, ScreenShotPath, rono, excelsave_path);
		}
	}
		
	private static void listFilesForFolder(File fileEntry)
	{
		
	}

	

}
