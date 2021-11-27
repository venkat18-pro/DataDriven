package org.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {

	public static void main(String[] args) throws IOException {
		
		File exeLoc = new File("C:\\Users\\ELCOT\\eclipse-workspace\\NewOne\\Excel\\UserInfo.xlsx");
		FileInputStream stream = new FileInputStream(exeLoc);
		Workbook w=new XSSFWorkbook(stream); 
		Sheet sheet = w.getSheet("Sheet1");
	
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				if (cellType == 1) {
					String sValue = cell.getStringCellValue();
					System.out.print(sValue+"  ");
				} else {
					if (DateUtil.isCellDateFormatted(cell)) {
						Date dateCellValue = cell.getDateCellValue();
						SimpleDateFormat std = new SimpleDateFormat("dd/MMM/yyyy");
						String dValue = std.format(dateCellValue);
						System.out.print(dValue+" ");
					} else {
						double numericCellValue = cell.getNumericCellValue();
						long num = (long)numericCellValue;
						String nValue = String.valueOf(num);
						System.out.print(nValue+" ");
					}
				}
			}
			System.out.println();
		}
	}

}
