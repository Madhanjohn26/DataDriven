package com.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataRead {

	//particular data
	private static void particular_Data() throws Exception {
		File f = new File("C:\\Users\\PC\\eclipse-workspace\\DataDriven\\User_Details.xlsx");
		FileInputStream fis =  new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		Row row = sheetAt.getRow(2);
		Cell cell = row.getCell(1);
		CellType cellType = cell.getCellType();
		
		if (cellType.equals(cellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		}
		else if (cellType.equals(CellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			int value = (int) numericCellValue;
			System.out.println(value);
			
		}
	}
//all data
	private static void all_Data() throws Exception {
		File f = new File("C:\\Users\\PC\\eclipse-workspace\\DataDriven\\User_Details.xlsx");
		FileInputStream fis =  new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		int rowSize = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < rowSize; i++) {
			Row row = sheetAt.getRow(i);
			int cellsize = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cellsize; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.print(stringCellValue+" ");
				}else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.print(value);
					
				}
			
			}
			System.out.println();
		}
	}
	
	
	//particular rows
	private static void particular_Row() throws Exception {
		File f = new File("C:\\Users\\PC\\eclipse-workspace\\DataDriven\\User_Details.xlsx");
		FileInputStream fis =  new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		int rowSize = sheetAt.getPhysicalNumberOfRows();
		//for (int i = 0; i < rowSize; i++) {
			Row row = sheetAt.getRow(1);
			int cellsize = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cellsize; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.print(stringCellValue);
				}else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.println(value);
					
				}
	                                                                	
			}
		//}
	}
	//particular column
	private static void particular_Column() throws Exception {
		File f = new File("C:\\Users\\PC\\eclipse-workspace\\DataDriven\\User_Details.xlsx");
		FileInputStream fis =  new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		int rowSize = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < rowSize; i++) {
			Row row = sheetAt.getRow(i);
			int cellsize = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cellsize; j++) {
				if (j==1) {					
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				}else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.println(value);
					
				}
				}
			}
		}
	}
	public static void main(String[] args) throws Exception {
		System.out.println("Particular Data");
		particular_Data();
		System.out.println("All Data");
		all_Data();
		System.out.println("Particular row");
		particular_Row();
		System.out.println("Particular Column");
		particular_Column();
	}
}
