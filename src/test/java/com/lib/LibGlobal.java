package com.lib;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.w3c.dom.Element;
import io.github.bonigarcia.wdm.WebDriverManager;
public class LibGlobal {

	static WebDriver driver;
	public static void getDriver() {
	WebDriverManager.chromedriver().setup();
	driver=new ChromeDriver();
	}
	 public static void  maximize() {
   driver.manage().window().maximize();
	}
	 public static void  loadUrl(String url) {
		 driver.get(url);
	}
	 public static void type(WebElement element,String data) {
	 element.sendKeys(data);
	 
	 }
	public static WebElement findElementById(String attributeValue) {
		 WebElement element = driver.findElement(By.id(attributeValue));
		 return element;
	} 
	
	public static WebElement findElemenyByXPath(String xpath) {
		WebElement element = driver.findElement(By.xpath(xpath));
		return element;

	}
	 public static void sendKeys(WebElement element,String data) {
		 element.sendKeys(data);
	}
	 public static WebElement findElementByName(String attributeValue) {
		 WebElement element = driver.findElement(By.name(attributeValue));
		 return element;
	}
	 public static void click(WebElement element) {
		element.click();
	}
	 public static Alert switchToAlert() {
	 Alert alert=driver.switchTo().alert();
	 return alert;
	 }
	 public static void clickOk( Alert alert) {
		 alert.accept();
	 }
	 public static void clickCancel (Alert alert) {
		 alert.dismiss();
	 }
	 public static void alertText (Alert alert,String text) {
		 alert.sendKeys(text);
	 }
	 public static String alertText (Alert alert) {
	return alert.getText();
	 }
	 
	 
	public static void selectByValue(WebElement element,String value) {
		Select select=new Select(element);
		select.selectByValue(value);
		}
	public static void selectOptionByText(WebElement element,String value) {
		Select select=new Select(element);
		select.selectByVisibleText(value);
		}
      public static  void clear(WebElement element) {
		   element.clear();

	}
      
      
      public static  void closeWindow() {
		   driver.close();

	}
      
      
      public static  void closeAllWindow() {
		   driver.quit();

	}
	  public static String getAttribute(WebElement element) {
		  String attribute = element.getAttribute("value");
		  return attribute;

	}
	  
	  public static String getText(WebElement element) {
		  String text=element.getText();
		return text;

	}
	  
	  
	  
	  
	  
	  public static  String getData( String sheetname ,int rownum,int cellnum) throws IOException {
			
			String res=null;
			File file=new File("C:\\Users\\Guest\\eclipse-workspace\\FrameWork\\Excel\\Baseclass.xlsx");
			FileInputStream stream=new FileInputStream(file);
			Workbook workbook=new XSSFWorkbook(stream);
			Sheet sheet = (Sheet) workbook.getSheet(sheetname);
			Row row = sheet.getRow(rownum);
			Cell cell = row.getCell(cellnum);
			CellType cellType = cell.getCellType();
			switch(cellType) {
			case STRING:
		    res= cell.getStringCellValue();
			break;
			case NUMERIC:
			if(DateUtil.isCellDateFormatted(cell)) {
				String a=new SimpleDateFormat("dd/mm/yyyy").format(cell.getDateCellValue());
				System.out.println(a);
			}
			else {
				double numericCellValue = cell.getNumericCellValue();
				BigDecimal b=new BigDecimal (numericCellValue);
				res = b.toString();
			}
				break;
			
				
			default:
				break;
			}
			
			return res;
			}
	  public static String UpdateData(String sheetname,int rownum,int cellnum,String olddata,String newdata) throws IOException{
			File file=new File("C:\\Users\\Guest\\eclipse-workspace\\FrameWork\\Excel\\Baseclass.xlsx");
			FileInputStream stream=new FileInputStream(file);
			Workbook workbook=new XSSFWorkbook(stream);
			Sheet sheet = workbook.getSheet(sheetname);
			Row row = sheet.getRow(rownum);
			Cell cell = row.getCell(cellnum);
			String stringCellValue = cell.getStringCellValue();
			if(stringCellValue.equals(olddata)) {
				cell.setCellValue(newdata);
			}
			FileOutputStream out=new FileOutputStream(file);
			workbook.write(out);
			return sheetname;
			
			
		}
		public static String insertValueInCell(String sheetName,int rownum,int cellnum,String data) throws IOException{
			File file=new File("C:\\Users\\Guest\\eclipse-workspace\\FrameWork\\Excel\\Baseclass.xlsx");
			FileInputStream stream=new FileInputStream(file);
			Workbook workbook=new XSSFWorkbook(stream);
			Sheet sheet = workbook.getSheet(sheetName);
			Row row = sheet.getRow(rownum);
			 Cell cell = row.createCell(cellnum);
			 cell.setCellValue(data);        
			 FileOutputStream out=new FileOutputStream(file);
			workbook.write(out);
			return sheetName;
	}
	  
		public static void getinsertValueInCell( WebElement element,String sheetName,int rownum,int cellnum) throws IOException {
			File file=new File("C:\\Users\\Guest\\eclipse-workspace\\FrameWork\\Excel\\Baseclass.xlsx");
			FileInputStream stream=new FileInputStream(file);
			Workbook workbook=new XSSFWorkbook(stream);
			Sheet sheet = workbook.getSheet(sheetName);
			Row row = sheet.getRow(rownum);
			 Cell cell = row.createCell(cellnum);
			String value = element.getAttribute("value");
			cell.setCellValue(value);
			 FileOutputStream out=new FileOutputStream(file);
				workbook.write(out);
				
		}
		public static void main(String[] args) {
			System.out.println("LIBRARY");
		}
		public static void main(String[] args) {
			System.out.println("===============================");
			System.out.println("BaseClass");
			System.out.println("***********************************");
		}

	    }
	
	
	
	
	
	
	
	
	

