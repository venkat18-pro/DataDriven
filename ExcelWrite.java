package org.demo;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws Exception {
		File excelFile = new File("C:\\Users\\ELCOT\\eclipse-workspace\\NewOne\\Excel\\ExcelWrite.xlsx"); 
		
		Workbook wBook = new XSSFWorkbook();
		Sheet sheet = wBook.createSheet("Data");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Name");
		
		FileOutputStream outStream = new FileOutputStream(excelFile);
		wBook.write(outStream);
		
		System.out.println("Excel File created..");
	}

}
