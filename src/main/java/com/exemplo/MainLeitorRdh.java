package com.exemplo;


import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainLeitorRdh {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		String RDH_PATH = "c:/temp/RDH_09OUT2021.xls";
		File file = new File(RDH_PATH);
		
		
		 
		//try(Workbook workbook = new XSSFWorkbook(file)){
		 try(Workbook workbook = WorkbookFactory.create(file)){
			
			// Busca a planilha
			Sheet sheet = workbook.getSheetAt(0);
			
			// Varre as linhas
			for(Row row: sheet) {
				for(Cell cell: row) {
					System.out.print(cell + "\t");
				}
				System.out.println();
			}
		}
	}
	
	
}
