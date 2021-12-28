package com.Maven_Pro2;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {
//steps to get excel sheet
//step1: get a file
//step2:get file input steam
//step3:get work sheet
	public static void main(String[] args) throws IOException {
    File f=new File
    		("C:\\Users\\Rajabi\\eclipse-workspace\\Data_Driven\\demo.xlsx");
    FileInputStream fis=new FileInputStream(f);
    Workbook wkb=new XSSFWorkbook(fis);
//from one excel book we can store more than one sheet
//get particular sheet by using getSheetAt()
    Sheet sheetAt = wkb.getSheetAt(1);
//wkb.getSheet(name)-----> can use this method to get excel sheet by using their name
    int num_row = sheetAt.getPhysicalNumberOfRows();
    System.out.println("numer of rows"+num_row);
//use nested for loop to get row and column data
    for(int i=0;i<num_row;i++) {
    	Row row = sheetAt.getRow(i);
    	int num_cell = row.getPhysicalNumberOfCells();
        for(int j=0;j<num_cell;j++)
        {
        	Cell cell = row.getCell(j);
        	CellType cellType = cell.getCellType();
  //-------------sell value may be srting or integer--------------//
  //-------------use if else condition for getting int or string value------//
        	if(cellType.equals(cellType.STRING)) {
        		String stringCellValue = cell.getStringCellValue();
        		System.out.println(stringCellValue);
        	}
        	else if(cellType.equals(cellType.NUMERIC)) {
        		double numericCellValue = cell.getNumericCellValue();
        		//numericCellVlue return type is double
        		//so convert into int by using type casting
        		int int_Value=(int)numericCellValue;
        		System.out.println(int_Value);
        	}
        }
    }
	}
	
}
