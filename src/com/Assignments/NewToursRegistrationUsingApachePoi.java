package com.Assignments;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class NewToursRegistrationUsingApachePoi 
{
	FileInputStream fin=null;
	XSSFWorkbook workBook=null;
	XSSFSheet sheet=null;
	Row row;
	Cell cell;
	
	String firstName;
	String lastName;
	String phone;
	String email;
	String address;
	String city;
	String state;
	String postalCode;
	String country;
	String userName;
	String password;
	String confirmedPass;
	
	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public String getLastName() {
		return lastName;
	}

	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public String getCity() {
		return city;
	}

	public void setCity(String city) {
		this.city = city;
	}

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}

	public String getPostalCode() {
		return postalCode;
	}

	public void setPostalCode(String postalCode) {
		this.postalCode = postalCode;
	}

	public String getCountry() {
		return country;
	}

	public void setCountry(String country) {
		this.country = country;
	}

	public String getUserName() {
		return userName;
	}

	public void setUserName(String userName) {
		this.userName = userName;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	public String getConfirmedPass() {
		return confirmedPass;
	}

	public void setConfirmedPass(String confirmedPass) {
		this.confirmedPass = confirmedPass;
	}
	
	public void setRegistrationData()
	{
		try
		{
			workBook=new XSSFWorkbook(new File("C:/Users/HP/Desktop/OrangeHRM.xlsx"));
			sheet=workBook.getSheet("NewToursRegistration");
	
			setFirstName(sheet.getRow(1).getCell(0).getStringCellValue()); 
			setLastName(sheet.getRow(1).getCell(1).getStringCellValue());
			cell=sheet.getRow(1).getCell(2);
			//converting numeric cell values to string avoid numeric cell exception
			if(cell.CELL_TYPE_NUMERIC==Cell.CELL_TYPE_NUMERIC)
			{
				cell.setCellType(Cell.CELL_TYPE_STRING);
			}
			setPhone(cell.getStringCellValue());
			setEmail(sheet.getRow(1).getCell(3).getStringCellValue());
			setAddress(sheet.getRow(1).getCell(4).getStringCellValue());
			setCity(sheet.getRow(1).getCell(5).getStringCellValue());
			setState(sheet.getRow(1).getCell(6).getStringCellValue());
			cell=sheet.getRow(1).getCell(7);
			//converting numeric cell values to string avoid numeric cell exception
			if(cell.CELL_TYPE_NUMERIC==Cell.CELL_TYPE_NUMERIC)
			{
				cell.setCellType(Cell.CELL_TYPE_STRING);
			}
			setPostalCode(cell.getStringCellValue());
			setCountry(sheet.getRow(1).getCell(8).getStringCellValue());
			setUserName(sheet.getRow(1).getCell(9).getStringCellValue());
			setPassword(sheet.getRow(1).getCell(10).getStringCellValue());
			setConfirmedPass(sheet.getRow(1).getCell(11).getStringCellValue());
			registration();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public void registration()
	{
		WebDriver driver=null;
		String url="http://newtours.demoaut.com/";
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		
		driver=new ChromeDriver();
		driver.get(url);
	
		//<a href="mercuryregister.php">REGISTER</a>
		WebElement register=driver.findElement(By.linkText("REGISTER"));
		
		String hrefAttribute = register.getAttribute("href");
		register.click();
		
		String actualUrl=driver.getCurrentUrl();
		
		if(actualUrl.equals(hrefAttribute))
		{
			driver.findElement(By.name("firstName")).sendKeys(getFirstName());
			driver.findElement(By.name("lastName")).sendKeys(getLastName());
			driver.findElement(By.name("phone")).sendKeys(getPhone());
			driver.findElement(By.name("userName")).sendKeys(getEmail());
			driver.findElement(By.name("address1")).sendKeys(getAddress());
			driver.findElement(By.name("city")).sendKeys(getCity());
			driver.findElement(By.name("state")).sendKeys(getState());
			driver.findElement(By.name("postalCode")).sendKeys(getPostalCode());
			Select select=new Select(driver.findElement(By.name("country")));
			select.selectByVisibleText(getCountry());
			driver.findElement(By.name("email")).sendKeys(getEmail());
			driver.findElement(By.name("password")).sendKeys(getPassword());
			driver.findElement(By.name("confirmPassword")).sendKeys(getConfirmedPass());
			String registerUrl="http://newtours.demoaut.com/create_account_success.php";
			
			driver.findElement(By.name("register")).click();
			if(registerUrl.equals(driver.getCurrentUrl()))
			{
				System.out.println("User registration successfull - Pass");
			}
			else
			{
				System.out.println("User registration failed - Fail");
			}
			
		}
		else
		{
			System.out.println("Failed to open Register page - Fail");
		}
		driver.close();
	}

	public static void main(String[] args)
	{
		new NewToursRegistrationUsingApachePoi().setRegistrationData();
	}
}
