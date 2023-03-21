// File contains all common functions can be used across any application


package Lib;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.*;
import Lib.GlobalVariables;
public class CommonFunctions {
	//public static WebDriver driver;
	//public static Properties OR;
	//public static XSSFWorkbook workbook;
	//public static XSSFSheet spreadsheet,spreadsheet1; 
	//public static XSSFRow row1,row;
	//public static Map < String, Object[] > summary,Detailsummary; 
	//public static String Filename =GlobalVariables.ExcelLog;
	//public static FileOutputStream out;
	//public static int rowid1,rowid;
	public CommonFunctions(){
			
	}
	public static String now(String dateFormat) {
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
		return sdf.format(cal.getTime());

	}
//	openBrowser() method used for opening the browser
	public void openBrowser() throws IOException
	{
		Runtime.getRuntime().exec("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255");
		GlobalVariables.driver = new FirefoxDriver();
	
	}
	
//	closeBrowser() Method used to close all browser opened during execution
	public void closeBrowser() 
	{
		GlobalVariables.driver.quit();
	}

// Navigate() method used to open a URL	
    public void Navigate(String inputurl) throws IOException
    {
    	GlobalVariables.driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    	GlobalVariables.driver.get(inputurl);
    	GlobalVariables.driver.manage().window().maximize();
    	GlobalVariables.driver.manage().deleteAllCookies();
    }

    // getObject method used to get an object with given property Returns "NULL" if object not found
    
	public static WebElement getObject(String locator,String element) throws IOException
	{
		
		int i;
		WebElement obj=null;
		switch (locator) {
        case "id": 
        	i=GlobalVariables.driver.findElements(By.id(element)).size();
        	//System.out.println("i="+i);
        	if(i==0){
	        	//System.out.println("Element is not present");
	        	obj=null;
	        	break;
        	}
        	else{
	        	//System.out.println("Element is present");
	        	obj=GlobalVariables.driver.findElement(By.id(element));
	        	break;
        	}
        case "xpath": 
        	 i=GlobalVariables.driver.findElements(By.xpath(element)).size();
        	//System.out.println("i="+i);
        	if(i==0){
	        	//System.out.println("Element is not present");
	        	obj=null;
	        	break;
        	}
        	else{
	        	//System.out.println("Element is present");
	        	obj=GlobalVariables.driver.findElement(By.xpath(element));
	        	break;
        	}
        case "name": 
	         
	       	 i=GlobalVariables.driver.findElements(By.name(element)).size();
	       	//System.out.println("i="+i);
	       	if(i==0){
		       	//System.out.println("Element is not present");
		       	obj=null;
		       	break;
	       	}
	       	else{
		       	//System.out.println("Element is present");
		       	obj=GlobalVariables.driver.findElement(By.name(element));
		        break;
	       	}
        case "linkText": 
	         
	       	 i=GlobalVariables.driver.findElements(By.linkText(element)).size();
	       	//System.out.println("i="+i);
	       	if(i==0){
		       	//System.out.println("Element is not present");
		       	obj=null;
		       	break;
	       	}
	       	else{
		       	//System.out.println("Element is present");
		       	obj=GlobalVariables.driver.findElement(By.linkText(element));
		        break;
	       	}
	       	
		}
		return obj;
	}
	
	
//	isVisible() method checks if given object is displayed
	public int isVisible(WebElement element){
		try{
			if(element.isDisplayed()){
				 return 1;
			}else
				return 0;
		}catch(Throwable t){
			// report error
			return 0;

		}
	}
	
// enterText() method enters text in a edit box	
	 public int enterText(WebElement element,String text) throws NoSuchElementException {
		try{
		 if(element.isDisplayed()){
			element.sendKeys(text);
			return 1;
		 }
		 else{
			//System.out.println("element not found ");
		 	return 0;
		 }
	}catch(Throwable t){
		// report error
		return 0;
	}
 }
	 
// Click() Method Clicks the object	 
	 public  int click(WebElement element){
			try{
				if(element.isDisplayed()){
					element.click();
				    return 1;
				}else
					return 0;
			}catch(Throwable t){
			// report error
			return 0;
			}
		}
	 
	 public int ClickRadiobtn(String name,int Option){
			try{
				 List<WebElement> radios = GlobalVariables.driver.findElements(By.name(name));
				 
				    if (Option > 0 && Option <= radios.size()) {
				    	elementHighlight(radios.get(Option - 1));
				        radios.get(Option - 1).click();
				        return 1;
				        
				}else
					return 0;
			}catch(Throwable t){
			// report error
			return 0;
			}
			
		}
	 
	 
// Select() Method select a value from the list 	 
		 public  int Select(WebElement element,String Text){
				try{
					Select list = new Select(element);
					
					int n =list.getOptions().size();
					for (int i=0;i<=n;i++){
						String str =list.getOptions().get(i).getText();
						if (str.equalsIgnoreCase(Text)){
							list.selectByIndex(i);
							break;
						}
					}
					//elementHighlight(list);
					//list.selectByValue(Text);
					return 1;
				}catch(Throwable t){
				// report error
				return 0;
				}
			}

	 //	 getText() method gets innertext of an object
	 public  String getText(WebElement element){
			try{
				if(element.isDisplayed()){
				     return element.getText();
				}else
					return null;
			}catch(Throwable t){
				// report error
				return null;
			}
		}
		
// store screenshots
		public  void takeScreenShot(String filePath) {
			File scrFile = ((TakesScreenshot)GlobalVariables.driver).getScreenshotAs(OutputType.FILE);
			try {
				FileUtils.copyFile(scrFile, new File(filePath));
			} catch (IOException e) {
				e.printStackTrace();
			}	   
		}


// Takes screenshots according to conditions
		public void Screenshot(String path,String condition,int RESULT){
			//System.out.println("Path : "+path+"condition : " +condition + " RESULT :"+ RESULT);
			if (condition.equalsIgnoreCase("ALL"))
				takeScreenShot(path);
			if (condition.equalsIgnoreCase("FAIL")&& (RESULT==0) )
				takeScreenShot(path);
			if (condition.equalsIgnoreCase("PASS")&& (RESULT==1) )
				takeScreenShot(path);
				
		}
// Highlight the element
		public void elementHighlight(WebElement element) {
			for (int i = 0; i < 2; i++) {
				JavascriptExecutor js = (JavascriptExecutor) GlobalVariables.driver;
				js.executeScript(
						"arguments[0].setAttribute('style', arguments[1]);",
						element, "color: red; border: 3px solid red;");
				js.executeScript(
						"arguments[0].setAttribute('style', arguments[1]);",
						element, "");
			}
		}

}