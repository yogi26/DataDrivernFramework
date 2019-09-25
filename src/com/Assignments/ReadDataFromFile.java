package com.Assignments;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromFile {

	public static void main(String[] args) 
	{
		FileInputStream fin;
		XSSFWorkbook workBook;
		XSSFSheet sheet;
		Row row;
		Cell cell;
		
		try
		{
			fin=new FileInputStream(new File("F:/eclipse/TestNG Framework/OrangeHRM.xlsx"));
			workBook=new XSSFWorkbook(fin);
			sheet=workBook.getSheet("Registration");
			row=sheet.getRow(1);
			cell=row.getCell(0);
			String data=cell.getStringCellValue();
			System.out.println(data);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

}
