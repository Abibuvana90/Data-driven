package com.Maven_Pro2;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadParticularDataUsingLoop {

	public static void main(String[] args) throws Throwable{
		File f=new File("C:\\Users\\Rajabi\\eclipse-workspace\\Data_Driven\\demo.xlsx");
	    FileInputStream fis=new FileInputStream(f);
	    Workbook wb=new XSSFWorkbook(fis);
	    Sheet sheet = wb.getSheet("username");
	    //to get row count
	    int rows_count = sheet.getPhysicalNumberOfRows();
	    System.out.println("number of rows "+rows_count);
//***************to read particular  data****************************//
//to get data from second row first column
//so change i=1;j=0 in excel row and col starts from 0
//change conditions in  both for loops [note:we dont give = so give next num to exit loop]
	    for(int i=1;i<2;i++) {
	    	Row row= sheet.getRow(i);
	    	int col_num = row.getPhysicalNumberOfCells();
	    	for(int j=0;j<1;j++) {
	    		Cell cell = row.getCell(j);
	    		CellType type_cell = cell.getCellType();
			    if(type_cell.equals(type_cell.STRING)) {
			    	String stringCellValue = cell.getStringCellValue();
			        System.out.print(stringCellValue);
			    }
			    else if(type_cell.equals(type_cell.NUMERIC)) {
			    	double numericCellValue = cell.getNumericCellValue();
			    	//note: above return type is in double so want to convert into in
			    	int numericCellValue1=(int)numericCellValue;
			    	System.out.print(numericCellValue1);
				}  
	    	}
	    }
	   
	    
		}
	}


