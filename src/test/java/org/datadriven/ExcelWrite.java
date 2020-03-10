package org.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {
		
		
		File f = new File("I:\\\\Aravinsami\\\\sami codes\\\\FirstMaven\\\\Excel\\\\Book1.xlsx");
		
		FileInputStream Is= new FileInputStream(f);
		
		Workbook w= new XSSFWorkbook();
		
		Sheet s = w.createSheet("Aravind");
		
		Row r = s.createRow(0);
		
		Cell c = r.createCell(0);
		
		c.setCellValue("Selenium");
		
		FileOutputStream Os= new FileOutputStream(f);
		
		w.write(Os);
		
		
	}
}
