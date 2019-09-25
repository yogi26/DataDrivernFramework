package com.Assignments;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.io.FileHandler;

import com.excel.utility.Xls_Reader;

public class OrangeHrmLoginUsingApachePoi 
{
	WebDriver driver=null;
	Xls_Reader reader=null;
	String url="http://localhost:8081/orangehrm/symfony/web/index.php/auth/login.html";
	
	FileInputStream fin=null;
	FileOutputStream fout=null;
	XSSFWorkbook workBook=null;
	XSSFSheet sheet=null;
	Row row=null;
	Cell cell=null;
	
	public void openBrowser()
	{
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		driver=new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(15, TimeUnit.SECONDS);
		driver.get(url);
	}
	public void login(String userName,String password)
	{
		try
		{
			Cell writeCell;
			
			driver.findElement(By.xpath("//input[@id='txtUsername']")).sendKeys(userName);
			driver.findElement(By.xpath("//input[@id='txtPassword']")).sendKeys(password);
			driver.findElement(By.xpath("//input[@id='btnLogin']")).click();
			
			//Validation for OrangeHRM homepage weather its open or not
			String dashboardUrl="http://localhost:8081/orangehrm/symfony/web/index.php/dashboard";
			if(dashboardUrl.equals(driver.getCurrentUrl()))
			{
				//written in 2nd cell of every row because Result cell is present at 2nd location in excel sheet
				writeCell=row.createCell(2);
				writeCell.setCellValue("Pass");
			}
			else
			{
				
				writeCell=row.createCell(2);
				writeCell.setCellValue("Fail");
				
				//below code is for taking screenshot of every fail test cases
				File file=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
				
				//Have set the name for every image using cell data so this will be uniquely identified 
				FileHandler.copy(file, new File("C:\\Users\\HP\\Desktop\\Images\\"+userName+"-"+password+".png"));
			}
			//write data in excel sheet
			fout=new FileOutputStream(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
			workBook.write(fout);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public void closeBrowser()
	{
		driver.quit();
	}
	public void getLoginData()
	{
		try 
		{
			fin=new FileInputStream(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
			workBook=new XSSFWorkbook(fin);
			sheet=workBook.getSheet("LoginModule");
			int rowCount=sheet.getLastRowNum();
			//checking weather rows are available or not in sheet
			if(rowCount!=0)
			{
				// row index started from 1st because at 0th row I gave cell heading name as Username, Password, Result
				
				for(int rowNum=1;rowNum<=rowCount;rowNum++)
				{
					row=sheet.getRow(rowNum);
					String userName=null;
					String password=null;
					
					for(int col=0;col<row.getLastCellNum();col++)
					{
						// condition for getting username and password separately
						if(col==0)
						{
							userName=row.getCell(col).getStringCellValue();
						}
						if(col==1)
						{
							password=row.getCell(col).getStringCellValue();
						}
					}
					openBrowser();
					System.out.println(userName+password);
					login(userName, password);
					closeBrowser();
				}
			}
			else
			{
				System.out.println("No rows found in given sheet");
			}
			
			
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
	public static void main(String[] args) 
	{
		new OrangeHrmLoginUsingApachePoi().getLoginData();
	}
}
