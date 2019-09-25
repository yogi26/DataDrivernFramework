package com.datadriven.test;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.excel.utility.Xls_Reader;

public class LoginTest
{
	public WebDriver driver=null;
	public Xls_Reader reader=null;
	String url="http://localhost:8081/orangehrm/symfony/web/index.php/auth/login.html";
	public void login()
	{
		reader=new Xls_Reader("F:\\eclipse\\DataDrivenFramework\\src\\com\\testdata\\OrangeHRM.xlsx");
		String userName=reader.getCellData("LoginModule", 0, 2);
		String password=reader.getCellData("LoginModule", 1, 2);
		
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(15, TimeUnit.SECONDS);
		
		driver.get(url);
		driver.findElement(By.xpath("//input[@id='txtUsername']")).sendKeys(userName);
		driver.findElement(By.xpath("//input[@id='txtPassword']")).sendKeys(password);
		driver.findElement(By.xpath("//input[@id='btnLogin']")).click();
	}
	public static void main(String[] args) 
	{
		new LoginTest().login();
	}

}
