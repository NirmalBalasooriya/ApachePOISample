package com.sample;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public static void main(String[] args) {
		try {
			String excelFilePath = "src/main/java/com/sample/Orders.xlsx";		 
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			System.out.println(excelFilePath);
			Workbook workbook = getRelevantWorkbook(inputStream, excelFilePath);
			
			Sheet firstSheet = workbook.getSheetAt(0);
	        Iterator<Row> iterator = firstSheet.iterator();
	         
	        while (iterator.hasNext()) {
	            Row nextRow = iterator.next();
	            Iterator<Cell> cellIterator = nextRow.cellIterator();
	            while (cellIterator.hasNext()) {
	                Cell cell = cellIterator.next();
	                switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					default:
						break;
					}
	                System.out.print(" ");
	            }
	            System.out.println();
	        }
	         
	        workbook.close();
	        inputStream.close();
			
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private static Workbook getRelevantWorkbook(FileInputStream inputStream, String excelFilePath) throws IOException
	{
	    Workbook workbook = null;
	 
	    if (excelFilePath.endsWith("xls")) {
	        workbook = new HSSFWorkbook(inputStream);
	    } else if (excelFilePath.endsWith("xlsx")) {
	        workbook = new XSSFWorkbook(inputStream);
	    } else {
	        throw new IllegalArgumentException("Incorrect file format");
	    }
	 
	    return workbook;
	}

}
