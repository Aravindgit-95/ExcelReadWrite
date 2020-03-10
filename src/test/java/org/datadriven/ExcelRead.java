package org.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("I:\\Aravinsami\\sami codes\\FirstMaven\\Excel\\Book1.xlsx");
		
		FileInputStream Is= new FileInputStream(f);
		
		Workbook w= new XSSFWorkbook(Is);
		
		Sheet s = w.getSheet("Sheet1");
		
	
		
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			
			Row r = s.getRow(i);
			
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				
				Cell c = r.getCell(j);
				
				CellType type = c.getCellType();
				
				if (type==CellType.STRING) {
					
					String value = c.getStringCellValue();
					
					System.out.println(value);
					
					
				}else if (type==CellType.NUMERIC) {
					
					if (DateUtil.isCellDateFormatted(c)) {
						
						Date d = c.getDateCellValue();
						
					SimpleDateFormat sd= new SimpleDateFormat("dd-mm-yy");
					String value = sd.format(d);
					System.out.println(value);
						
					}
					
					
					else {
						
						double dd = c.getNumericCellValue();
						
						long l= (long)dd;
						String value = String.valueOf(l);
						
						System.out.println(value);
					}
					
				}
				
				
					
					
					
				
				
			}
			
			
			
		}
		
	}

}
