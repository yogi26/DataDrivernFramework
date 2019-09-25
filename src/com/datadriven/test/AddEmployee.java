package com.datadriven.test;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.excel.utility.Xls_Reader;

public class AddEmployee
{
	public WebDriver driver=null;
	public void addEmployee() 
	{
		String sheetName="addEmployee";
		String xlsFilePath="F:\\eclipse\\DataDrivenFramework\\src\\com\\testdata\\OrangeHRM.xlsx";
		String url="http://localhost:8081/orangehrm/symfony/web/index.php/auth/login.html";
		
		Xls_Reader reader=new Xls_Reader(xlsFilePath);
		
		//System.out.println(firstName+" "+middleName+" "+lastName+" "+empId);
		
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		driver=new ChromeDriver();
		driver.get(url);
		driver.manage().window().maximize();
		driver.findElement(By.xpath("//input[@id='txtUsername']")).sendKeys("Admin");
		driver.findElement(By.xpath("//input[@id='txtPassword']")).sendKeys("Admin");
		driver.findElement(By.xpath("//input[@id='btnLogin']")).click();
		driver.findElement(By.linkText("PIM")).click();
		driver.findElement(By.id("btnAdd")).click();
		int rowCount=reader.getRowCount(sheetName);
		for(int i=2;i<=rowCount;i++)
		{
			String firstName=reader.getCellData(sheetName, 0, i);
			String middleName=reader.getCellData(sheetName, 1, i);
			String lastName=reader.getCellData(sheetName, 2, i);
			String empId=reader.getCellData(sheetName, 3, i);
			
			driver.findElement(By.id("firstName")).clear();
			driver.findElement(By.id("middleName")).clear();
			driver.findElement(By.id("lastName")).clear();
			
			driver.findElement(By.id("firstName")).sendKeys(firstName);
			driver.findElement(By.id("middleName")).sendKeys(middleName);
			driver.findElement(By.id("lastName")).sendKeys(lastName);
			driver.findElement(By.id("employeeId")).sendKeys(empId);
			driver.findElement(By.id("btnSave")).click();
			
			driver.navigate().back();
		}
		driver.close();
	}
	public static void main(String[] args)
	{
		new AddEmployee().addEmployee();
	}

}
