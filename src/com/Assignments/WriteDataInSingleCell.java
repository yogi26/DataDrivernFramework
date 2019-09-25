package com.Assignments;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataInSingleCell
{
	public static void main(String[] args)
	{
		FileInputStream fin;
		FileOutputStream fout;
		XSSFWorkbook workBook;
		XSSFSheet sheet;
		Row row;
		Cell cell;
		try
		{
			fin=new FileInputStream(new File("F:/eclipse/TestNG Framework/OrangeHRM.xlsx"));
			
			workBook=new XSSFWorkbook(fin);
			sheet=workBook.getSheet("Sheet3");
			row=sheet.createRow(0);
			cell=row.createCell(0);
			cell.setCellValue("Yogesh");
			cell=row.createCell(1);
			cell.setCellValue("Giri");
			fout=new FileOutputStream("F:/eclipse/TestNG Framework/OrangeHRM.xlsx");
			workBook.write(fout);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
}
