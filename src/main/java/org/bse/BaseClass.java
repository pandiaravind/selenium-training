package org.bse;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
public static  WebDriver driver;
	
	public static void launchBrowser() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		}
	
	public static void windowMaximize() {
	driver.manage().window().maximize();
}
	public static void launchURL(String url) {
		driver.get(url);
}
	public static void pageTitle() {
		String title = driver.getTitle();
System.out.println(title);
	}
	public static void pageURL() {
		String Url = driver.getCurrentUrl();
		System.out.println(Url);
}
	public static void passText(String txt,WebElement ele) {
		ele.sendKeys(txt);
}
	public static void closeentirebrowser() {
		driver.quit();
}
	public static void clickbtn(WebElement ele) {
	ele.click();	
		}
	public static void screenshot(String imgName)throws IOException {
	TakesScreenshot ts = (TakesScreenshot) driver;  	
    File image = ts.getScreenshotAs(OutputType.FILE);
    File f = new File("Location+ imgName.png");
    FileUtils.copyFile(image, f);
	}
	public static Actions a;
	
	public static void movethecursor(WebElement targetwebelement) {
		a = new Actions (driver);
		a.moveToElement(targetwebelement).perform();
		}
	public static void drangdrop(WebElement dragWebElement,WebElement dropElement) {
	a = new Actions(driver);
	a.dragAndDrop(dragWebElement, dropElement).perform();
    }
	public static  JavascriptExecutor js;
	
	public static void scrollthepage(WebElement tarWebEle) {
	js = (JavascriptExecutor)driver;
	js.executeScript("arguments[0].scrollIntoView(false)",tarWebEle);
	}
	public static void scroll(WebElement element) {
		js = (JavascriptExecutor)driver;
		js.executeScript("arguments[0].scrollIntoView(true)", element);
		}
	public static void excelread(String SheetName,int rowName,int cellName)throws IOException {
		File f = new File("C:\\\\Users\\\\ELCOT\\\\eclipse-workspace\\\\FirstProgram\\\\Excel\\\\newXLfile.xlsx");
		FileInputStream fils = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fils);
		Sheet mysheet = wb.getSheet("data");
		Row r= mysheet.getRow(rowName);
		Cell c = r.getCell(cellName);
		int cellType = c.getCellType();
		
		String value =" ";
		if (cellType==1) {
			String Value2 = c.getStringCellValue();
			}
		else if (DateUtil.isCellDateFormatted(c)) {
			Date dd = c.getDateCellValue();
			SimpleDateFormat s = new SimpleDateFormat(value);
			String value1 = s.format(dd);
		}
		
		else {
			double d = c.getNumericCellValue();
			long l = (long)d;
			String valueOf = String.valueOf(l);
		}
		}
	public static void creatnewExcell(int rowNum,int cellNum,String writeData)throws IOException {
		File f = new File("C:\\\\Users\\\\ELCOT\\\\eclipse-workspace\\\\FirstProgram\\\\Excel\\\\newXLfile.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet newSheet =w.createSheet("Datas");
		Row newRow = newSheet.createRow(rowNum);
		Cell newCell =newRow.createCell(cellNum);
		newCell.setCellValue(writeData);
		FileOutputStream fos =new FileOutputStream(f);
		w.write(fos);
	}
	public static void creatcell(int getRow,int creCell,String newData)throws IOException {
		File f = new File("C:\\\\Users\\\\ELCOT\\\\eclipse-workspace\\\\FirstProgram\\\\Excel\\\\newXLfile.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet s = wb.getSheet("Datas");
		Row r = s.getRow(getRow);
		Cell c = r.createCell(creCell);
		c.setCellValue(newData);
		FileOutputStream fos =new FileOutputStream(f);
		wb.write(fos);
	}
	public static void creatrow(int creRow,int creCell,String newData) throws IOException {
		File f = new File("C:\\Users\\ELCOT\\eclipse-workspace\\FirstProgram\\Excel\\newXLfile.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet s = wb.getSheet("Datas");
		Row r = s.createRow(creRow);
		Cell c = r.createCell(creCell);
		c.setCellValue(newData);
		FileOutputStream fos =new FileOutputStream(f);
		wb.write(fos);
		
		
	}
	
	public static void updatedatatoperticularcell(int getTheRow,int getTheCell,String exisitingData,String writeNewData) throws IOException {
    File f = new File("C:\\\\Users\\\\ELCOT\\\\eclipse-workspace\\\\FirstProgram\\\\Excel\\\\newXLfile.xlsx");
    FileInputStream fis = new FileInputStream(f);
    Workbook wb = new XSSFWorkbook(fis);
	Sheet s = wb.getSheet("Datas");
	Row r = s.getRow(getTheRow);
	Cell c = r.getCell(getTheCell);
	String str = c.getStringCellValue();
	if (str.equals(exisitingData)) {
		c.setCellValue(writeNewData);
		
	}
	FileOutputStream fos = new FileOutputStream(f);
	wb.write(fos);
	

	}

}
