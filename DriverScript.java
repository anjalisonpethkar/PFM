// This is main script to drive the automation suite

package Script;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.data.general.DefaultPieDataset;

import Lib.keyWords;
import Lib.GlobalVariables;
import Lib.CommonFunctions;

import Lib.Xls_Reader;

public class DriverScript {

public static keyWords keywrds;
public static CommonFunctions CF;

//Functional Driver Declarations
public static Xls_Reader FunctionalDriver;

public static Xls_Reader XLSLog;
public int KeywordStart = 5;

public static int result;
public static String screenshotfile;
public static String TeststartTime,TestendTime,KewrdstartTime,KewrdendTime;
public static Long lTeststartTime, lTestendTime,lKewrdstartTime,lKewrdendTime,TotalKwrdExTime,TotalTestExTime;
public static String strTotalTCtime,strTotalKWtime;

public DriverScript() throws IOException
{
	keywrds = new keyWords();
	CF = new CommonFunctions();
}

	public static void main(String[] args) throws IOException, InterruptedException, NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
     DriverScript drvscrpt= new DriverScript();
     
     String startTime = CommonFunctions.now("dd.MMMMM.yyyy hh.mm.ss aaa");
     drvscrpt.CreateHTMLFile(startTime,"EndTime",GlobalVariables.gTotalTC,GlobalVariables.gTotalTCpass,GlobalVariables.gTotalTCfail);
     drvscrpt.CreateXLlogfile();
     drvscrpt.Execute_Flow();
     
     drvscrpt.drawchart(GlobalVariables.ExcelLog, GlobalVariables.Lastrow +4);
     
     

     }


 public void Execute_Flow() throws IOException, InterruptedException, NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException 
 {
	 String TestNo , Testcase_Name1,Result1,ExecutedOn1,Total_Run_Time1,strResult1 = null,Result2;
	 String TotalKWtime,TransTime1,TransTime2,Screenshot_Name1;
	 int KW =2, Krw=0;
	 int Resrow;   
	 int TotalTC=0,TotalPassTC=0,TotalFailTC=0;
	 String TChttmlFilename;
	FunctionalDriver =  new Xls_Reader(System.getProperty("user.dir")+"\\src\\datatables\\FunctionalDriver.xls");
	GlobalVariables.IamGlobal=5;
	GlobalVariables.Testcase_Master_row=1;
	 for (GlobalVariables.Testcase_Master_row =2;GlobalVariables.Testcase_Master_row<= FunctionalDriver.getRowCount("TestCase_Master");GlobalVariables.Testcase_Master_row++){
		 if(FunctionalDriver.getColumnValue("TestCase_Master",GlobalVariables.TobeExecuted, GlobalVariables.Testcase_Master_row).equals("Y")){
			SetCurrentID();			 	
			//System.out.println(GlobalVariables.currenttestCaseID+" "+GlobalVariables.currenttestCaseName);
			TChttmlFilename = System.getProperty("user.dir")+"\\src\\HTMLReports\\"+GlobalVariables.currenttestCaseID+GlobalVariables.timestamp+".html";
			CreateTestcaseHTMLFile(TChttmlFilename, GlobalVariables.currenttestCaseID);
			for (GlobalVariables.Business_Flow_row=2;GlobalVariables.Business_Flow_row<= FunctionalDriver.getRowCount("Business_Flow");GlobalVariables.Business_Flow_row++){
				if(BusinessFlowMatched() ==1){
					TeststartTime=CommonFunctions.now("dd.MMMMM.yyyy hh.mm.ss aaa");
					//System.out.println("TeststartTime " +TeststartTime);
					lTeststartTime=new Date().getTime(); 
					//System.out.println("lTeststartTime " +lTeststartTime);
					Krw=0;
					//System.out.println(Business_Flow_row+" ISflowmatch "+BusinessFlowMatched()); 
					for(int methodIndex = KeywordStart;methodIndex<FunctionalDriver.getColumnCount("Business_Flow",GlobalVariables.Business_Flow_row-1);methodIndex++){
						
						String keyword = FunctionalDriver.getColumnValue("Business_Flow", methodIndex, GlobalVariables.Business_Flow_row);
						//System.out.println("Keyword name " +keyword);
						KewrdstartTime=CommonFunctions.now("dd.MMMMM.yyyy hh.mm.ss aaa");
						TransTime1=KewrdstartTime;
						//System.out.println("KewrdstartTime " +KewrdstartTime);
						lKewrdstartTime=new Date().getTime(); 
						//System.out.println("lKewrdstartTime " +lKewrdstartTime);
						if (keyword.equals("End"))
							break;		
						else{
							java.lang.reflect.Method method;
							 method= keywrds.getClass().getMethod(keyword);
							 result = (int) method.invoke(keywrds);
							 
							 if (result==0){
								 	strResult1 ="FAIL";
									
								// System.out.println("strResult1 " +strResult1);
								 //System.out.println("Keyword " +keyword+" Failed");
								 screenshotfile=GlobalVariables.ScreenshotPath+GlobalVariables.currenttestCaseID+keyword+GlobalVariables.timestamp+".jpg";
							 	 //System.out.println("Name of screenshot " +screenshotfile);
							 	 CF.Screenshot(screenshotfile, GlobalVariables.ScreenshotFor, result);
							 	 CF.closeBrowser();
							 	 KewrdendTime =CommonFunctions.now("dd.MMMMM.yyyy hh.mm.ss aaa");
								 //System.out.println("KewrdendTime " +KewrdendTime);
								 lKewrdendTime = new Date().getTime();
								 //System.out.println("lKewrdendTime " +lKewrdendTime);
								 TotalKwrdExTime = lKewrdendTime - lKewrdstartTime;
 								//System.out.println("TotalKwrdExTime " +TotalKwrdExTime);
								strTotalKWtime= ConvertMillisconds(TotalKwrdExTime);
								//System.out.println("Keyword: "+keyword+" strTotalKWtime " +strTotalKWtime);
								
								//System.out.println("strResultDetail " +strResult);
								TestNo =GlobalVariables.currenttestCaseID; Testcase_Name1=GlobalVariables.currenttestCaseName;Result2=strResult1;
								TotalKWtime=strTotalKWtime;TransTime2=KewrdendTime;
								Screenshot_Name1 = screenshotfile;
								//System.out.println("Result2 " +Result2);
								Krw++;
								AddDetailrow(KW,TestNo,keyword,Result2,TotalKWtime, TransTime1,TransTime2,Screenshot_Name1);
								AppendTestcaseHTMLFile(TChttmlFilename,Integer.toString(Krw),keyword,Result2,TransTime1,TransTime2,TotalKWtime,Screenshot_Name1);
							    KW++;
			
							 	 break;
							 }
							 else
									strResult1 ="PASS";
							 //System.out.println("Result of Keyword " +keyword+" is"+result);
							 //System.out.println("strResult1 " +strResult1);
							 
							 if (keyword.equalsIgnoreCase("CloseallBrowsers"))
								 CF.Screenshot(screenshotfile, "None", result);
							 else	 
							 {
								 screenshotfile=GlobalVariables.ScreenshotPath+GlobalVariables.currenttestCaseID+keyword+GlobalVariables.timestamp+".jpg";
							 	// System.out.println("Name of screenshot " +screenshotfile);
							 	 CF.Screenshot(screenshotfile, GlobalVariables.ScreenshotFor, result);
							 } 
							}
						KewrdendTime =CommonFunctions.now("dd.MMMMM.yyyy hh.mm.ss aaa");
						//System.out.println("KewrdendTime " +KewrdendTime);
						lKewrdendTime = new Date().getTime();
						//System.out.println("lKewrdendTime " +lKewrdendTime);
						TotalKwrdExTime = lKewrdendTime - lKewrdstartTime;
						
						//System.out.println("TotalKwrdExTime " +TotalKwrdExTime);
						strTotalKWtime= ConvertMillisconds(TotalKwrdExTime);
						//System.out.println("Keyword: "+keyword+" strTotalKWtime " +strTotalKWtime);
						
						//System.out.println("strResultDetail " +strResult);
						TestNo =GlobalVariables.currenttestCaseID; Testcase_Name1=GlobalVariables.currenttestCaseName;Result2=strResult1;
						TotalKWtime=strTotalKWtime;TransTime2=KewrdendTime;
						Screenshot_Name1 = screenshotfile;
						if(keyword.equalsIgnoreCase("CloseallBrowsers"))
							Screenshot_Name1="";
						//System.out.println("Result2 " +Result2);
						Krw++;
						AddDetailrow(KW,TestNo,keyword,Result2,TotalKWtime, TransTime1,TransTime2,Screenshot_Name1);
						AppendTestcaseHTMLFile(TChttmlFilename,Integer.toString(Krw),keyword,Result2,TransTime1,TransTime2,TotalKWtime,Screenshot_Name1);
					    KW++;
					}
					TestendTime=CommonFunctions.now("dd.MMMMM.yyyy hh.mm.ss aaa");
					//System.out.println("TestendTime " +TestendTime);
					lTestendTime = new Date().getTime();
					//System.out.println("lTestendTime " +lTestendTime);
					TotalTestExTime = lTestendTime - lTeststartTime;
					//System.out.println("TotalTestExTime " +TotalTestExTime);
					strTotalTCtime= ConvertMillisconds(TotalTestExTime);
					//System.out.println("strTotalTCtime " +strTotalTCtime);
					if(result==1) {
					} else {
					}
					//System.out.println("strResult " +strResult);
					Resrow =GlobalVariables.Testcase_Master_row;TestNo =GlobalVariables.currenttestCaseID; Testcase_Name1=GlobalVariables.currenttestCaseName;Result1=strResult1;ExecutedOn1=TeststartTime;Total_Run_Time1=strTotalTCtime;
					if(result==1)
						TotalPassTC++;
					else
						TotalFailTC++;
				    AddSummaryrow(Resrow,TestNo,Testcase_Name1,Result1,ExecutedOn1, Total_Run_Time1);
				    AppendMaineHTMLFile(GlobalVariables.currenttestCaseID,Integer.toString(Resrow-1),Result1,ExecutedOn1,Total_Run_Time1,TChttmlFilename);
				    TotalTC++;
				    GlobalVariables.Lastrow = Resrow;
				    
				}
			}
    	 }
	 }
	 //System.out.println("TotalTC " +TotalTC+"PassTC "+TotalPassTC+"FailTC "+TotalFailTC);
	 
	 GlobalVariables.gTotalTC=TotalTC; GlobalVariables.gTotalTCpass=TotalPassTC;GlobalVariables.gTotalTCfail=TotalFailTC;
	 XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	 XLSLog.setCellData1("Summary ", "Testcase_Name", GlobalVariables.Lastrow +3, "Total Tcs tobe executed");
	 XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	 XLSLog.setCellData1("Summary ", "Result", GlobalVariables.Lastrow +3, Integer.toString(TotalTC));
	 XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	 XLSLog.setCellData1("Summary ", "Testcase_Name", GlobalVariables.Lastrow +4, "Total Tcs Passed");
	 XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	 XLSLog.setCellData1("Summary ", "Result", GlobalVariables.Lastrow +4, Integer.toString(TotalPassTC));
	 XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	 XLSLog.setCellData1("Summary ", "Testcase_Name", GlobalVariables.Lastrow +5, "Total Tcs Failed");
	 XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	 XLSLog.setCellData1("Summary ", "Result", GlobalVariables.Lastrow +5, Integer.toString(TotalFailTC));
	 
	 updatePassFail(CommonFunctions.now("dd.MMMMM.yyyy hh.mm.ss aaa"),Integer.toString(TotalPassTC),Integer.toString(TotalFailTC),Integer.toString(TotalTC));
	 	 	 	
 }
 
 // Method to set Current Testcase ID
 public void SetCurrentID(){
	 GlobalVariables.currentBSID = FunctionalDriver.getColumnValue("TestCase_Master", GlobalVariables.BS_ID, GlobalVariables.Testcase_Master_row); 
	 GlobalVariables.currentTSID = FunctionalDriver.getColumnValue("TestCase_Master", GlobalVariables.TS_ID, GlobalVariables.Testcase_Master_row);
	 GlobalVariables.currentTCID = FunctionalDriver.getColumnValue("TestCase_Master", GlobalVariables.TC_ID, GlobalVariables.Testcase_Master_row);
	 GlobalVariables.currenttestCaseID = GlobalVariables.currentBSID+"-"+GlobalVariables.currentTSID+"-"+GlobalVariables.currentTCID;
	 GlobalVariables.currenttestCaseName = FunctionalDriver.getColumnValue("TestCase_Master", GlobalVariables.TestcaseName, GlobalVariables.Testcase_Master_row);
 }
 
 // Method to to get matching Business flow for a testcase
 public int BusinessFlowMatched(){
	 GlobalVariables.BF_BSID = FunctionalDriver.getColumnValue("Business_Flow", GlobalVariables.BS_ID, GlobalVariables.Business_Flow_row); 
	 GlobalVariables.BF_TS_ID = FunctionalDriver.getColumnValue("Business_Flow", GlobalVariables.TS_ID, GlobalVariables.Business_Flow_row); 
	 GlobalVariables.BF_TC_ID = FunctionalDriver.getColumnValue("Business_Flow", GlobalVariables.TC_ID, GlobalVariables.Business_Flow_row);
		if (GlobalVariables.BF_BSID.equals(GlobalVariables.currentBSID) && GlobalVariables.BF_TS_ID.equals(GlobalVariables.currentTSID) && GlobalVariables.BF_TC_ID.equals(GlobalVariables.currentTCID))
			return 1;
		else
			return 0;
 }
 public String ConvertMillisconds(Long ms){
	 
	 //System.out.println("ms " +ms);
	 Long MilliSec = ms % 1000;
	 Long Sec = ms/1000;
	 Long min = Sec/60;
	 Sec = Sec %60;
	 Long Hr = min/60;
	 min= min%60;

	 if (Hr != 0){
	 String time = Hr+" Hours "+min +" Minutes " +Sec +" Seconds " + MilliSec +" MilliSeconds";
	 //System.out.println("time " +time);
	 return time;
	 }
	 
	 if ((Hr == 0) && (min!=0)){
		 String time = min +" Minutes " +Sec +" Seconds " + MilliSec +" MilliSeconds";
		// System.out.println("time " +time);
		 return time;
		 }
	 if ((Hr == 0) && (min==0)){
		 String time = Sec +" Seconds " + MilliSec +" MilliSeconds";
		 //System.out.println("time " +time);
		 return time;
		 }
	 return "";
	  
 }
 

public void AddSummaryrow(int i,String Test_no,String Testcase_Name,String Result,String ExecutedOn,String Total_Run_Time)
 {
 	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
 	XLSLog.setCellData("Summary ", "Test_No", i, Test_no);
 	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Summary ", "Testcase_Name", i, Testcase_Name);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Summary ", "Result", i, Result);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Summary ", "ExecutedOn", i, ExecutedOn);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Summary ", "Total_Run_Time", i, Total_Run_Time);
 	
 }
public void AddDetailrow(int i,String Test_Name,String Keyword_Name,String Resultn,String Total_Run_Time,String Transaction_Start_Time,String Transaction_End_Time,String Screenshot_Name)
{
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Details ", "Test_Name", i, Test_Name);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Details ", "Keyword_Name", i, Keyword_Name);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Details ", "Result", i, Resultn);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Details ", "Total_Run_Time", i, Total_Run_Time);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Details ", "Transaction_Start_Time", i, Transaction_Start_Time);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Details ", "Transaction_End_Time", i, Transaction_End_Time);
	XLSLog= new Xls_Reader(GlobalVariables.ExcelLog);
	XLSLog.setCellData("Details ", "Screenshot_Name", i, Screenshot_Name);
}
public boolean AddTotalSummary(int row, int Col,String TC) throws IOException
{
	FileInputStream fis = new FileInputStream(GlobalVariables.ExcelLog); 
	HSSFWorkbook workbook = new HSSFWorkbook(fis);

	if(row<=0){
		workbook.close();
		return false;
	}
	
	int index = workbook.getSheetIndex("Summary ");
	
	HSSFSheet sheet = workbook.getSheetAt(index);
		
	HSSFCell cell; 	
	cell=sheet.createRow(row).createCell(Col);
	 
	CellStyle style = workbook.createCellStyle();
	Font hSSFFont = workbook.createFont();
    hSSFFont.setFontHeightInPoints((short) 12);
    hSSFFont.setFontName("Times New Roman");
    hSSFFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
    hSSFFont.setColor(HSSFColor.YELLOW.index);
    style.setFont(hSSFFont);
    style.setBorderBottom((short) 1);
    style.setBorderLeft((short) 1);
    style.setBorderRight((short) 1);
    style.setBorderTop((short) 1);
	style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
	style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	style.setWrapText(true);
	cell.setCellStyle(style);
	
    
    FileOutputStream fileOut = new FileOutputStream(GlobalVariables.ExcelLog);

	workbook.write(fileOut);

    fileOut.close();	
    workbook.close();
    return true;


}


public void CreateXLlogfile() throws IOException{
	HSSFWorkbook workbook = new HSSFWorkbook(); 
    //Create a blank sheet
    HSSFSheet spreadsheet = workbook.createSheet("Summary ");
    //Create row object
    HSSFRow row;
 
    //This data needs to be written (Object[])
    Map < String, Object[] > Summary =      new TreeMap < String, Object[] >();
    Summary.put( "1", new Object[] {    "Test_No", "Testcase_Name", "Result","ExecutedOn","Total_Run_Time" });
    
    //Iterate over data and write to sheet
    Set < String > keyid = Summary.keySet();
    int rowid = 0;
    for (String key : keyid)
    {
    	CellStyle style = workbook.createCellStyle();
    	HSSFFont hSSFFont = workbook.createFont();
        hSSFFont.setFontHeightInPoints((short) 12);
        hSSFFont.setFontName("Times New Roman");
        hSSFFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        hSSFFont.setColor(HSSFColor.YELLOW.index);
        style.setFont(hSSFFont);
    	style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
    	style.setFillPattern(CellStyle.SOLID_FOREGROUND);
    	
    	
       row = spreadsheet.createRow(rowid++);
       //row.setHeightInPoints(12.75f);
      // row.setRowStyle(style);
       Object [] objectArr = Summary.get(key);
       int cellid = 0;
       for (Object obj : objectArr)
       {
          Cell cell = row.createCell(cellid++);
          cell.setCellStyle(style);
          cell.setCellValue((String)obj);
       }
    }
  //Create a blank sheet
    HSSFSheet spreadsheet1 = workbook.createSheet("Details ");
    //Create row object
    HSSFRow row1;
    //This data needs to be written (Object[])
    Map < String, Object[] > Details =      new TreeMap < String, Object[] >();
    Details.put( "1", new Object[] {    "Test_Name", "Keyword_Name", "Result","Total_Run_Time","Transaction_Start_Time","Transaction_End_Time","Screenshot_Name" });
    
    //Iterate over data and write to sheet
    Set < String > keyid1 = Details.keySet();
    int rowid1 = 0;
    for (String key : keyid1)
    { 
    	CellStyle style = workbook.createCellStyle();
    	HSSFFont hSSFFont = workbook.createFont();
    	hSSFFont.setFontHeightInPoints((short) 12);
    	hSSFFont.setFontName("Times New Roman");
    	hSSFFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
    	hSSFFont.setColor(HSSFColor.YELLOW.index);
    	style.setFont(hSSFFont);
    	style.setFillForegroundColor(IndexedColors.DARK_GREEN.getIndex());
    	style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	
       row1 = spreadsheet1.createRow(rowid1++);
       Object [] objectArr = Details.get(key);
       int cellid = 0;
       for (Object obj : objectArr)
       {
    	   
          Cell cell = row1.createCell(cellid++);
          cell.setCellStyle(style);
          cell.setCellValue((String)obj);
       }
    }
    //Write the workbook in file system
    String Filename =GlobalVariables.ExcelLog;
    FileOutputStream out = new FileOutputStream( 
    new File(Filename));
    workbook.write(out);
    out.close();
    workbook.close();
    

}

public void CreateHTMLFile(String startTime,String EndTime,int totalTC,int PassTC,int failTC) throws UnknownHostException
{
	
	FileWriter fstream =null;
	BufferedWriter out =null;
	
	InetAddress IP=InetAddress.getLocalHost();
	System.out.println("IP of my system is := "+IP.getHostAddress());
	String RUN_DATE = CommonFunctions.now("dd.MMMMM.yyyy").toString();
	try{
		// Create file 

		fstream = new FileWriter(GlobalVariables.HTMLlog);
		out = new BufferedWriter(fstream);
		out.newLine();

		out.write("<html>\n");
		out.write("<HEAD>\n");
		out.write(" <TITLE>Automation Test Results</TITLE>\n");
		out.write("</HEAD>\n");

		out.write("<body>\n");
		out.write("<h4 align=center><FONT COLOR=660066 FACE=Times New Roman SIZE=6><b><u> Automation Test Results</u></b></h4>\n");
		out.write("<table  border=1 cellspacing=1 cellpadding=1 >\n");
		out.write("<tr>\n");

		out.write("<h4> <FONT COLOR=660000 FACE=Times New Roman SIZE=4.5> <u>Test Details Executed on :  "+IP+" </u></h4>\n");
		out.write("<td width=150 align=left bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2.75><b>Run Date</b></td>\n");
		out.write("<td width=150 align=left><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2.75><b>"+RUN_DATE+"</b></td>\n");
		out.write("</tr>\n");
		out.write("<tr>\n");

		out.write("<td width=150 align=left bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2.75><b>Run StartTime</b></td>\n");

		out.write("<td width=200 align=left><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2.75><b>"+startTime+"</b></td>\n");
		out.write("</tr>\n");
		out.write("<tr>\n");
		// out.newLine();   
		out.write("<td width=150 align= left  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2.75><b>END_TIME</b></td>\n");
		out.write("<td width=200 align= left ><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2.75><b>"+EndTime+"</b></td>\n");
		out.write("</tr>\n");
		out.write("<tr>\n");
		
		//out.write("<td width=150 align= left  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2.75><b>TotalTimeTaken</b></td>\n");
		//out.write("<td width=200 align= left ><FONT COLOR=#153E7E FACE= Arial  SIZE=2.75><b>TIME_TAKEN</b></td>\n");
		//out.write("</tr>\n");
		//out.write("<tr>\n");
		//  out.newLine();

		out.write("<td width=150 align= left  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2.75><b>Passed</b></td>\n");
		out.write("<td width=150 align= left ><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2.75><b>passed</b></td>\n");
		out.write("</tr>\n");
		out.write("<tr>\n");

		out.write("<td width=150 align= left  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2.75><b>Failed</b></td>\n");
		out.write("<td width=150 align= left ><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2.75><b>failed</b></td>\n");
		out.write("</tr>\n");
		out.write("</tr>\n");
		
		out.write("<td width=150 align= left  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2.75><b>Total No.Testcases</b></td>\n");
		out.write("<td width=150 align= left ><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2.75><b>count</b></td>\n");
		out.write("</tr>\n");
		
		out.write("<table  border=1 cellspacing=1 cellpadding=1 width=100%>\n");
		out.write("<tr>\n");
		out.write("<td width=20% align= center  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2><b>TestCaseID</b></td>\n");
		out.write("<td width=20% align= center  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2><b>Test Case Name</b></td>\n");
		out.write("<td width=10% align= center  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2><b>Status</b></td>\n");
		out.write("<td width=20% align= center  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2><b>Executed ON</b></td>\n");
		out.write("<td width=30% align= center  bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE= Times New Roman  SIZE=2><b>Testcase Execution Time</b></td>\n");
		
		out.write("</tr>\n");
		out.write("</table>\n");


		//Close the output stream
		out.close();


	}catch (Exception e){//Catch exception if any
		System.err.println("Error: " + e.getMessage());
	}finally{

		fstream=null;
		out=null;
	}
}

public void CreateTestcaseHTMLFile(String TCNamefile,String TCName)
{
	FileWriter fstream =null;
	BufferedWriter out =null;
	
	
	try{
		// Create file 

		fstream = new FileWriter(TCNamefile);
		out = new BufferedWriter(fstream);
		out.newLine();

		out.write("<html>\n");
		out.write("<HEAD>\n");
		out.write(" <TITLE> <FONT COLOR=660000 FACE=Arial SIZE=4.5>"+TCName+ "Execution Test Results</TITLE>\n");
		out.write("</HEAD>\n");

		out.write("<body>\n");
		out.write("</body>\n");
		out.write("<h4> <FONT COLOR=660000 FACE=Times New Roman SIZE=4.5> Detailed Report :</h4>");
		out.write("<table  border=1 cellspacing=1    cellpadding=1 width=100%>");
		out.write("<tr> ");
		out.write("<td align=center width=5%  align=center bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2><b>Step/Row#</b></td>");
		out.write("<td align=center width=10% align=center bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2><b>Keyword</b></td>");
		out.write("<td align=center width=10% align=center bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2><b>Result</b></td>");
		out.write("<td align=center width=15% align=center bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2><b>Keyword Start</b></td>");
		out.write("<td align=center width=15% align=center bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2><b>Keyword End</b></td>");
		out.write("<td align=center width=15% align=center bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2><b>Keyword Execution Time</b></td>");
		
		out.write("<td align=center width=40% align=center bgcolor=#153E7E><FONT COLOR=#E0E0E0 FACE=Times New Roman SIZE=2><b>Screen Shot</b></td>");
		out.write("</tr>");
		out.write("</table>");
			//Close the output stream
		out.close();
		

	}catch (Exception e){//Catch exception if any
		System.err.println("Error: " + e.getMessage());
	}finally{

		fstream=null;
		out=null;
	}
}
public void AppendTestcaseHTMLFile(String TCNamefile,String Step,String keyword,String Result,String KWstartTm,String KWEndTm,String KWTTime,String Screenshot){
	 FileWriter fstream =null;
	BufferedWriter out =null;
 
	try{
		fstream = new FileWriter(TCNamefile,true);
		out = new BufferedWriter(fstream);
		out.write("<table  border=1 cellspacing=1    cellpadding=1 width=100%>");
		out.write("<tr> ");
		out.write("<td align=center width=5%  align=center bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2>"+Step+"</td>");
		out.write("<td align=center width=10% align=center bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2>"+keyword+"</td>");
		out.write("<td align=center width=10% align=center bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2>"+Result+"</td>");
		out.write("<td align=center width=15% align=center bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2>"+KWstartTm+"</td>");
		out.write("<td align=center width=15% align=center bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2>"+KWEndTm+"</td>");
		out.write("<td align=center width=15% align=center bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2>"+KWTTime+"</td>");
		out.write("<td align=center width=40% align=center bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE=Times New Roman SIZE=2><a href="+Screenshot+">"+Screenshot+"</td>");
		out.write("</tr>");
		out.write("</table>");
		
		
		out.close();
	}catch(Exception e){
		System.err.println("Error: " + e.getMessage());
	}finally{

		fstream=null;
		out=null;
	}
   
  }
public void AppendMaineHTMLFile(String TestCaseID,String TestCaseNm,String Status,String Executed,String TCTTime,String Filenm){
	 FileWriter fstream =null;
	BufferedWriter out =null;

	try{
		fstream = new FileWriter(GlobalVariables.HTMLlog,true);
		out = new BufferedWriter(fstream);
		out.write("<table  border=1 cellspacing=1 cellpadding=1 width=100%>\n");
		out.write("<tr>\n");
		out.write("<td width=20% align= center  bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2>"+TestCaseNm+"</td>\n");
		out.write("<td width=20% align= center  bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2><a href="+Filenm+">"+TestCaseID+"</td>\n");
		out.write("<td width=10% align= center  bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2>"+Status+"</td>\n");
		out.write("<td width=20% align= center  bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2>"+Executed+"</td>\n");
		out.write("<td width=30% align= center  bgcolor=#E0E0E0><FONT COLOR=#153E7E FACE= Times New Roman  SIZE=2>"+TCTTime+"</td>\n");
		
		out.write("</tr>\n");
		out.write("</table>\n");
		
		out.close();
	}catch(Exception e){
		System.err.println("Error: " + e.getMessage());
	}finally{

		fstream=null;
		out=null;
	}
  
 }



public static void updatePassFail(String EndTime,String pass,String fail,String count)
{
	StringBuffer buf = new StringBuffer();
	try{
		// Open the file that is the first 
		// command line parameter
		FileInputStream fstream = new FileInputStream(GlobalVariables.HTMLlog);
		// Get the object of DataInputStream
		DataInputStream in = new DataInputStream(fstream);
		BufferedReader br = new BufferedReader(new InputStreamReader(in));
		String strLine;



		//Read File Line By Line

		while ((strLine = br.readLine()) != null)   {
			if(strLine.indexOf("EndTime") !=-1){
				strLine=strLine.replace("EndTime", EndTime);
			}
			if(strLine.indexOf("passed") !=-1){
				strLine=strLine.replace("passed", pass);
			}
			if(strLine.indexOf("failed") !=-1){
				strLine=strLine.replace("failed", fail);
			}
			
			if(strLine.indexOf("count") !=-1){
				strLine=strLine.replace("count", count);
			}
			
			buf.append(strLine);


		}
		//Close the input stream
		in.close();
		//System.out.println(buf);
		FileOutputStream fos=new FileOutputStream(GlobalVariables.HTMLlog);
		DataOutputStream   output = new DataOutputStream (fos);	 
		output.writeBytes(buf.toString());
		fos.close();

	}catch (Exception e){//Catch exception if any
		System.err.println("Error: " + e.getMessage());
	}
}

public  void drawchart(String Filename, int r) throws IOException{
	 FileInputStream chart_file_input = new FileInputStream(new File(Filename));
       /* Read chart data from XLSX Workbook */
       HSSFWorkbook my_workbook = new HSSFWorkbook(chart_file_input);
       /* Read worksheet that has pie chart input data information */
       HSSFSheet my_sheet = my_workbook.getSheetAt(0);
       /* Create JFreeChart object that will hold the Pie Chart Data */
       DefaultPieDataset my_pie_chart_data = new DefaultPieDataset();

       
       /* Loop through worksheet data and populate Pie Chart Dataset */
       String chart_label="a";
       Number chart_data=0;            

    
       /* Add data to the data set */  
       Row row1,row2;
       
       row1=my_sheet.getRow(r-1);
       Cell cell= row1.getCell(1);
       chart_label= cell.getStringCellValue();
       System.out.println("chart_label "+chart_label);
        Cell cell1= row1.getCell(2);
       chart_data=Integer.parseInt(cell1.getStringCellValue());
       System.out.println("chart_data "+chart_data);
        my_pie_chart_data.setValue(chart_label,chart_data);
       row2=my_sheet.getRow(r);
       Cell cell2= row2.getCell(1);
       chart_label=  cell2.getStringCellValue();
       System.out.println("chart_label "+chart_label);
         Cell cell3= row2.getCell(2);
       chart_data=Integer.parseInt(cell3.getStringCellValue());
       System.out.println("chart_data "+chart_data);
       my_pie_chart_data.setValue(chart_label,chart_data);
                     
       /* Create a logical chart object with the chart data collected */
       JFreeChart myPieChart=ChartFactory.createPieChart("Test Execution Result",my_pie_chart_data,true,true,false);
       /* Specify the height and width of the Pie Chart */
       int width=440; /* Width of the chart */
       int height=380; /* Height of the chart */
       float quality=1; /* Quality factor */
       /* We don't want to create an intermediate file. So, we create a byte array output stream 
       and byte array input stream
       And we pass the chart data directly to input stream through this */             
       /* Write chart as JPG to Output Stream */
       ByteArrayOutputStream chart_out = new ByteArrayOutputStream();          
       ChartUtilities.writeChartAsJPEG(chart_out,quality,myPieChart,width,height);
       /* We now read from the output stream and frame the input chart data */
       /* We don't need InputStream, as it is required only to convert the output chart to byte array */
       /* We can directly use toByteArray() method to get the data in bytes */
       /* Add picture to workbook */
       int my_picture_id = my_workbook.addPicture(chart_out.toByteArray(), Workbook.PICTURE_TYPE_JPEG);                
       /* Close the output stream */
       chart_out.close();
       /* Create the drawing container */
       HSSFPatriarch drawing = my_sheet.createDrawingPatriarch();
       /* Create an anchor point */
       ClientAnchor my_anchor = new HSSFClientAnchor();
       /* Define top left corner, and we can resize picture suitable from there */
       my_anchor.setCol1(4);
       my_anchor.setRow1(r-1);
       /* Invoke createPicture and pass the anchor point and ID */
       HSSFPicture  my_picture = drawing.createPicture(my_anchor, my_picture_id);
       /* Call resize method, which resizes the image */
       my_picture.resize();
       /* Close the FileInputStream */
       chart_file_input.close();               
       /* Write Pie Chart back to the XLSX file */
       FileOutputStream out = new FileOutputStream(new File(Filename));
       my_workbook.write(out);
       out.close(); 
       my_workbook.close();

}


}





 
 
