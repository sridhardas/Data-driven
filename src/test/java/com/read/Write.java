package com.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public  class Write {
	public static void write() throws Throwable {
		File f=new File("‪‪C:\\Users\\sridhar\\OneDrive\\Documents\\SRI.xlsx");
FileInputStream fis=new FileInputStream(f);
Workbook w=new XSSFWorkbook(fis);
Sheet createSheet = w.createSheet("User_Details");
Row createRow = createSheet.createRow(0);
Cell createCell = createRow.createCell(0);
createCell.setCellValue("UserName");
w.getSheet("User_Details").getRow(0).createCell(1).setCellValue("Password");
w.getSheet("User_Details").createRow(1).createCell(0).setCellValue("smith");
w.getSheet("User_Details").getRow(1).createCell(1).setCellValue("abc123");
w.getSheet("User_Details").createRow(2).createCell(0).setCellValue("hulk");
FileOutputStream fos=new FileOutputStream(f);
w.write(fos);
w.close();
System.out.println("done");
	}

	public static void main(String[] args) throws Throwable {
write();
	}

}
