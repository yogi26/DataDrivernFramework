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

public class ExportTimeAndDate
{
	WebDriver driver=null;
	String url="https://www.timeanddate.com/worldclock//";
	
	FileInputStream fin=null;
	FileOutputStream fout=null;
	XSSFWorkbook workBook=null;
	XSSFSheet sheet=null;
	Row row=null;
	Cell cell=null;
	
	public ExportTimeAndDate()
	{
		try
		{
			fin=new FileInputStream(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
			workBook=new XSSFWorkbook(fin);
			sheet=workBook.createSheet("Time and Dates of Countries");
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		openBrowser();
		exportDateAndTime();
		closeBrowser();
	}
	public void openBrowser()
	{
		System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
		driver=new ChromeDriver();
		driver.get(url);
		
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
		driver.manage().deleteAllCookies();
	}
	public void exportDateAndTime()
	{
		String expectedTitle="The World Clock — Worldwide";
		String actualTitle=driver.getTitle();
		System.out.println(actualTitle);
		if(actualTitle.equals(expectedTitle))
		{
			try
			{
				WebElement table=driver.findElement(By.xpath("/html/body/div[1]/div[6]/section[1]/div/section/div[1]/div/table/tbody"));
				List<WebElement> rows=table.findElements(By.tagName("tr"));
				
				for(int i=0;i<rows.size();i++)
				{
					row=sheet.createRow(i);
					List<WebElement> cols=rows.get(i).findElements(By.tagName("td"));
					for(int j=0;j<cols.size();j++)
					{
						cell=row.createCell(j);
						String myData=cols.get(j).getText();
						cell.setCellValue(myData);
					}
				}
				fout=new FileOutputStream(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
				workBook.write(fout);
			}
			catch(Exception e)
			{
				e.printStackTrace();
			}
		}
		else
		{
			System.out.println("Failed To Open The World Clock - Worldwide Page - Fail");
		}
		
	}
	public void closeBrowser()
	{
		driver.close();
	}
	public static void main(String[] args) 
	{
		new ExportTimeAndDate();
	}
}
