package com.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataWrite {
	
	public static void write_Data() throws Exception {
		File f = new File("C:\\Users\\PC\\Desktop\\ExcelDataWrite.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet createSheet = wb.createSheet("User_Data");
		Row createRow = createSheet.createRow(0);
		Cell createCell = createRow.createCell(0);
		createCell.setCellValue("UserName");
		wb.getSheet("User_Data").getRow(0).createCell(1).setCellValue("Password");
		wb.getSheet("User_Data").createRow(1).createCell(0).setCellValue("Madhan");
		wb.getSheet("User_Data").getRow(1).createCell(1).setCellValue("123");
		wb.getSheet("User_Data").createRow(2).createCell(0).setCellValue("Jairo");
		wb.getSheet("User_Data").getRow(2).createCell(1).setCellValue("456");
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		wb.close();
		System.out.println("Process Completed");
	}
	
	public static void main(String[] args) throws Throwable  {
		write_Data();
	}

}
