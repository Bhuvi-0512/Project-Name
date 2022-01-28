package org.setup;

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

public class ExcelRead {
	
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
//				System.out.println(cell);
				int cellType = cell.getCellType();
				System.out.println(cellType);
				
				
				
				if(cellType==1) {
					String Value = cell.getStringCellValue();
					System.out.println(Value);
				}
					
					else {
						if(DateUtil.isCellDateFormatted(cell)) {
							Date date = cell.getDateCellValue();
							SimpleDateFormat si = new SimpleDateFormat("dd/MM/yyy");
							String Value = si.format(date);
							System.out.println(Value);						
						}
						else {
							double phno = cell.getNumericCellValue();
							long l= (long) phno;
							String value = String.valueOf(l);
							System.out.println(value);
						}
						
						
						
					}				
				
			}
			
		}
				
	}
		
}


	
	
