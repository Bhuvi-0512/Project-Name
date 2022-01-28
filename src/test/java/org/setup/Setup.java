package org.setup;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Setup {
	
	public static void main(String[] args) throws IOException {		
		File loc = new File("C:\\Users\\HP\\eclipse-workspace\\org.setup\\src\\test\\resources\\Data Driven Task\\today task.xlsx");
				
		FileInputStream FI = new FileInputStream(loc);
		
		Workbook W = new XSSFWorkbook(FI);
		
		Sheet S = W.getSheet("Sheet1");
		
		Row R = S.getRow(3);
		
		Cell C = R.getCell(2);
		
		int AllRows = S.getPhysicalNumberOfRows();
		System.out.println(AllRows);
		
		int Cells = R.getPhysicalNumberOfCells();
		System.out.println(Cells);
		
		for(int i=0;i<AllRows;i++) {
			Row row = S.getRow(i);			
			for(int j=0;j<row.getPhysicalNumberOfCells();j++){
				Cell cell = row.getCell(j);
				System.out.println(cell);				
			}
			
		}
		
		
		
	
		
		
		
		
		
		
		
	}
	
	
	
	
	
	
	
	
}


	
	
