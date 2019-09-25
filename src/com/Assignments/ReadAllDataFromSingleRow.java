package com.Assignments;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadAllDataFromSingleRow 
{
	public static void main(String[] args)
	{
		FileInputStream fin;
		XSSFWorkbook workBook;
		XSSFSheet sheet;
		Row row;
		int rowCount;
		Cell cell;
		
		try
		{
			fin=new FileInputStream(new File("F:/eclipse/TestNG Framework/OrangeHRM.xlsx"));
			workBook=new XSSFWorkbook(fin);
			sheet=workBook.getSheet("Registration");
			rowCount=sheet.getLastRowNum();
			for(int i=1;i<=rowCount;i++)
			{
				row=sheet.getRow(i);
				cell=row.getCell(0);
				System.out.println(cell.getStringCellValue());
			}
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}

	}

}
