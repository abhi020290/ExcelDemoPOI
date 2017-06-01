package com.howtodoinjava.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo 
{
	@SuppressWarnings("null")
	public static void main(String[] args) 
	{
		
		
		List<Integer> status = new ArrayList<Integer>();
		List<String> contid =  new ArrayList<String>();
		
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		
		try
		{
			FileInputStream file = new FileInputStream(new File("c://temp/cleanup.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			
			while (rowIterator.hasNext()) 
			{
				
				Row row = rowIterator.next();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
						case Cell.CELL_TYPE_NUMERIC:
							/*System.out.print("value of id " +cell.getNumericCellValue() + "\t");*/
							status.add((int) cell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							/*System.out.print("Value od status " +cell.getStringCellValue() + "\t");*/
							contid.add(cell.getStringCellValue());
							break;
					}
				}
				System.out.println("");
				
			}
			
			System.out.println(status);
			System.out.println(contid);
			
			
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
}
