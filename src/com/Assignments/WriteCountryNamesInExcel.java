package com.Assignments;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class WriteCountryNamesInExcel 
{	
	WebDriver driver=null;
	String url="http://newtours.demoaut.com/";
	
	FileInputStream fin=null;
	FileOutputStream fout=null;
	XSSFWorkbook workBook=null;
	XSSFSheet sheet=null;
	Row row=null;
	Cell cell=null;
	
	public WriteCountryNamesInExcel()
	{
		try
		{
			fin=new FileInputStream(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
			workBook=new XSSFWorkbook(fin);
			sheet=workBook.createSheet("CountryNames");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		openBrowser();
		writeCountryNames();
		closeBrowser();
	}
	public void openBrowser()
	{
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		driver=new ChromeDriver();
		driver.get(url);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
	}
	public void writeCountryNames()
	{
		try
		{
			WebElement register=driver.findElement(By.linkText("REGISTER"));
			String registerURL=register.getAttribute("href");
			register.click();
			if(registerURL.equals(driver.getCurrentUrl()))
			{
				WebElement country=driver.findElement(By.name("country"));
				List<WebElement> countryList=country.findElements(By.tagName("option"));
				row=sheet.createRow(0);
				cell=row.createCell(0);
				cell.setCellValue("Country Names");
				for(int i=1;i<countryList.size();i++)
				{
					row=sheet.createRow(i);
					cell=row.createCell(0);
					cell.setCellValue(countryList.get(i).getText());
				}
				fout=new FileOutputStream(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
				workBook.write(fout);
			}
			else
			{
				System.out.println("Failed to open Register Page - Fail");
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public void closeBrowser()
	{
		driver.close();
	}
	public static void main(String[] args)
	{
		new WriteCountryNamesInExcel();
	}
}
