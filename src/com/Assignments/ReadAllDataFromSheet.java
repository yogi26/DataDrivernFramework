package com.Assignments;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAllDataFromSheet
{
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
			int rowCount=sheet.getLastRowNum();
			int cellCount;
			for(int rowNum=1; rowNum<=rowCount; rowNum++)
			{
				cellCount=0;
				row=sheet.getRow(rowNum);
				cellCount=row.getLastCellNum();
				for(int cellNum=0; cellNum<cellCount; cellNum++)
				{
					cell=row.getCell(cellNum);
					if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
					{
						cell.setCellType(Cell.CELL_TYPE_STRING);
					}
					System.out.print(cell.getStringCellValue()+" ");
				}
				System.out.println();
			}	
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

}
