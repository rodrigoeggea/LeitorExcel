package com.exemplo;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class Main {

	public static void main(String[] args) throws IOException {
		String fileLocation = "C:/temp/customers.xls";				
		FileInputStream file = new FileInputStream(new File(fileLocation));
		
		// Workbook workbook = new XSSFWorkbook(file);     // XLSX
		Workbook workbook = WorkbookFactory.create(file);  // XLS
		Sheet sheet = workbook.getSheetAt(0);
	
		for(Row row: sheet) {
			for(Cell cell: row) {
				System.out.print(cell + "\t");
			}
			System.out.println();
		}
		workbook.close();
	}
}
