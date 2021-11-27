package org.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUpdate {

	public static void main(String[] args) throws Exception {
		File excelFile = new File("C:\\Users\\ELCOT\\eclipse-workspace\\NewOne\\Excel\\ExcelWrite.xlsx");
		FileInputStream inStream = new FileInputStream(excelFile);
		
		Workbook wBook = new XSSFWorkbook(inStream);
		
		Sheet sheet = wBook.getSheet("Data");
		Row row = sheet.getRow(0);
		
		Cell cell = row.getCell(0);
		String sValue = cell.getStringCellValue();
		if(sValue.equals("Name")) {
			cell.setCellValue("Venkat");
		}
		
		FileOutputStream outStream = new FileOutputStream(excelFile);
		
		wBook.write(outStream);
		
		System.out.println("File Updated..");
		
		
	}

}
