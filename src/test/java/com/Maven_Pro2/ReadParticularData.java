package com.Maven_Pro2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadParticularData {

	public static void main(String[] args) throws Throwable {
    File f=new File("C:\\Users\\Rajabi\\eclipse-workspace\\Data_Driven\\demo.xlsx");
    FileInputStream fis=new FileInputStream(f);
    Workbook wb=new XSSFWorkbook(fis);
    Sheet sheetAt = wb.getSheetAt(0);
    //to get row count
    int rows_count = sheetAt.getPhysicalNumberOfRows();
    System.out.println("number of rows "+rows_count);
    //following two lines are used to retrieve particular data 
    //row num2: col num3
    Row row = sheetAt.getRow(2);
    Cell cell = row.getCell(3);
    
   CellType type_cell = cell.getCellType();
    if(type_cell.equals(type_cell.STRING)) {
    	String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
    }
    else if(type_cell.equals(type_cell.NUMERIC)) {
    	double numericCellValue = cell.getNumericCellValue();
    	//note: above return type is in double so want to convert into in
    	int numericCellValue1=(int)numericCellValue;
    	System.out.println(numericCellValue1);
	}
    
	}
}


