package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Sample {
	
	public static void main(String[] args) throws IOException {
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Sony\\eclipse-workspace\\Project\\driver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
		
		driver.get("https://www.aircanada.com/");
		
		driver.manage().window().maximize();
		
		String Ttle = driver.getTitle();
		
		System.out.println(Ttle);
		
		File loc = new File("C:\\Users\\Sony\\eclipse-workspace\\Project\\Excel\\TestData.xlsx");
		
		FileInputStream stream = new FileInputStream(loc);
		
		//workbook
		
		Workbook w = new XSSFWorkbook(stream);
		
		//sheet
		
		 org.apache.poi.ss.usermodel.Sheet s = w.getSheet("Data");
		 
		 //row
		 
		 Row r = s.getRow(1);
		 
		 //cell
		 
		 Cell c = r.getCell(1);
		 
		 System.out.println(c);
		 
		
		
		
	}
	

}
