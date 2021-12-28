package com.Maven_Pro2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws Throwable {
    File f=new File("C:\\Users\\Rajabi\\eclipse-workspace\\Data_Driven\\demo.xlsx");
    FileInputStream fis=new FileInputStream(f);
    Workbook wb=new XSSFWorkbook(fis);
    //to write data in excel the following steps needed
    //1.create a sheet using----->createSheet("name")
    //2.create a row--------->createRow(rownum)
    //3.create a cell-------->createCell(index)
    //4.set the value-------->setCellValue(boolean)
    //should select setCellValue(boolean) here boolean mandatory
    
    //create a sheet
       Sheet Sheet_name = wb.createSheet("WriteData");
   
       Row cr0 =Sheet_name.createRow(0);
       cr0.createCell(0).setCellValue("30");
       cr0.createCell(1).setCellValue("4");
       cr0.createCell(2).setCellValue("20");
       Row cr1 = Sheet_name.createRow(1);
       cr1.createCell(0).setCellValue("abi");
       cr1.createCell(1).setCellValue("mithu");
       cr1.createCell(2).setCellValue("tina");
       
    FileOutputStream fos=new FileOutputStream(f);
    wb.write(fos);
    wb.close();
    System.out.println("excel sheet created.....");
	}

}
