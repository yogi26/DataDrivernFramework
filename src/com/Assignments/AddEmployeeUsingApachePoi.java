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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class AddEmployeeUsingApachePoi 
{

	WebDriver driver=null;
	String url="http://localhost:8081/orangehrm/symfony/web/index.php/auth/login.html";
	FileInputStream fin=null;
	FileOutputStream fout=null;
	XSSFWorkbook workBook;
	XSSFSheet sheet;
	Row row;
	Cell cell;
	
	String firstName;
	String middleName;
	String lastName;
	String empId;
	String userName;
	String password;
	
	public String getUserName()
	{
		return userName;
	}

	public void setUserName(String userName)
	{
		this.userName = userName;
	}

	public String getPassword()
	{
		return password;
	}

	public void setPassword(String password)
	{
		this.password = password;
	}

	public String getFirstName()
	{
		return firstName;
	}

	public void setFirstName(String firstName)
	{
		this.firstName = firstName;
	}
	
	public String getMiddleName()
	{
		return middleName;
	}
	
	public void setMiddleName(String middleName)
	{
		this.middleName = middleName;
	}

	public String getLastName()
	{
		return lastName;
	}

	public void setLastName(String lastName)
	{
		this.lastName = lastName;
	}

	public String getEmpId()
	{
		return empId;
	}
	
	public void setEmpId(String empId)
	{
		this.empId = empId;
	}
	public void setData()
	{
		try
		{
			fin=new FileInputStream(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
			workBook=new XSSFWorkbook(fin);
			
			//setting up login data
			sheet=workBook.getSheet("Login");
			setUserName(sheet.getRow(0).getCell(0).getStringCellValue());
			setPassword(sheet.getRow(0).getCell(1).getStringCellValue());
			
			//setting up add employee data
			sheet=workBook.getSheet("AddEmployee");
			
			for(int rownum=1;rownum<=sheet.getLastRowNum();rownum++)
			{
				setFirstName(sheet.getRow(rownum).getCell(0).getStringCellValue());
				setMiddleName(sheet.getRow(rownum).getCell(1).getStringCellValue());
				setLastName(sheet.getRow(rownum).getCell(2).getStringCellValue());
				setEmpId(sheet.getRow(rownum).getCell(3).getStringCellValue());	
			}	
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}	
	}
	public void openBrowser()
	{
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		driver=new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(15, TimeUnit.SECONDS);
		driver.get(url);
		login();
	}
	public void login()
	{
		try
		{			
			driver.findElement(By.xpath("//input[@id='txtUsername']")).sendKeys(getUserName());
			driver.findElement(By.xpath("//input[@id='txtPassword']")).sendKeys(getPassword());
			driver.findElement(By.xpath("//input[@id='btnLogin']")).click();
			
			//Validation for OrangeHRM home page weather its open or not
			String dashboardUrl="http://localhost:8081/orangehrm/symfony/web/index.php/dashboard";
			if(dashboardUrl.equals(driver.getCurrentUrl()))
			{
				WebElement pim=driver.findElement(By.linkText("PIM"));
				Actions act=new Actions(driver);
				act.moveToElement(pim).perform();
				
				WebElement addEmpPage=driver.findElement(By.xpath("//*[@id='menu_pim_addEmployee']"));
				String addEmpPageLink =addEmpPage.getAttribute("href");
				addEmpPage.click();	
				
				if(addEmpPageLink.equals(driver.getCurrentUrl()))
				{
					addEmployee();					
				}
				else
				{
					System.out.println("Failed to open Add Empoyee - Fail");
				}
			}
			else
			{
				System.out.println("Failed to login - Fail");
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public void addEmployee()
	{
		System.out.println("Add Employee page opened - Pass");
		
		WebElement firstName=driver.findElement(By.xpath("//*[@id='firstName']"));
		firstName.sendKeys(getFirstName());
		
		WebElement middleName=driver.findElement(By.xpath("//*[@id='middleName']"));
		middleName.sendKeys(getMiddleName());
		
		WebElement lastname=driver.findElement(By.xpath("//*[@id='lastName']"));
		lastname.sendKeys(getLastName());
		
		WebElement empId=driver.findElement(By.xpath("//*[@id='employeeId']"));
		empId.clear();
		empId.sendKeys(getEmpId());
		driver.findElement(By.id("btnSave")).click();
		
		if(driver.getCurrentUrl().contains("pim/viewPersonalDetails"))
		{
			System.out.println("Employee Added Successfully - Pass");
			
			String empName=driver.findElement(By.xpath("//*[@id='profile-pic']/h1")).getText();
			String eId=driver.findElement(By.id("personal_txtEmployeeId")).getAttribute("value");
			
			String fullName=getFirstName()+" "+getMiddleName()+" "+getLastName();
			
			if((empName.equals(fullName))&&(eId.equals(getEmpId())))
			{
				System.out.println("Employee Name and Id validated - Pass");
			}
			else
			{
				System.out.println("Employee Name and Id validated - Fail");
			}
		}
		else
		{
			System.out.println("Failed to Add Employee - Fail");
		}
	}
	public void closeBrowser()
	{
		driver.close();
	}
	public static void main(String[] args)
	{
		AddEmployeeUsingApachePoi obj=new AddEmployeeUsingApachePoi();
		obj.setData();
		obj.openBrowser();
		obj.closeBrowser();	
	}

}
