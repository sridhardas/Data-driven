package com.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.XLSBUnsupportedException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read {
public static void particular_Data() throws Throwable {
	System.out.println("particular data is");
	File f=new File("C:\\Users\\sridhar\\OneDrive\\Desktop\\Datadrivenread.xlsx");
	FileInputStream fis=new FileInputStream(f);
	Workbook w=new XSSFWorkbook(fis);
	Sheet sheetAt = w.getSheetAt(0);
	Row row = sheetAt.getRow(2);
	Cell cell = row.getCell(1);
	CellType cellType = cell.getCellType();
	if (cellType.equals(CellType.STRING)) {
		String stringCellValue = cell.getStringCellValue();
		System.out.println(stringCellValue);
	}
	else if (cellType.equals(CellType.NUMERIC)) {
		double numericCellValue = cell.getNumericCellValue();
		int value=(int) numericCellValue;
		
		//int value=(int) numericCellValue;
		System.out.println(value);
	}
	
}

public static void allData() throws Throwable {
	System.out.println();
	System.out.println("all data is");
File f =new File("C:\\Users\\sridhar\\OneDrive\\Desktop\\Datadrivenread.xlsx");
	FileInputStream fis=new FileInputStream(f);
	Workbook w=new XSSFWorkbook(fis);
	Sheet sheetAt = w.getSheetAt(0);
	int numbersofRows = sheetAt.getPhysicalNumberOfRows();
	//System.out.println(Rows);
	for (int i = 0; i <numbersofRows; i++) {
		Row row = sheetAt.getRow(i);
		int NumberOfCells = row.getPhysicalNumberOfCells();
		for (int j = 0; j <NumberOfCells; j++) {
			Cell cell = row.getCell(j);
			CellType cellType = cell.getCellType();
			if (cellType.equals(CellType.STRING)) {
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
			}
			else if (cellType.equals(CellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				int value=(int) numericCellValue;
				System.out.println(value);
			}
		}
	}
	
}
public static void particular_Row() throws Throwable {
	System.out.println();
	System.out.println("particular row data");
File f =new File("C:\\Users\\sridhar\\eclipse-workspace\\DataDriven-Framework\\Datadrivenread.xlsx");
FileInputStream fis=new FileInputStream(f);
	Workbook w=new XSSFWorkbook(fis);
	Sheet sheetAt = w.getSheetAt(0);
	int numberOfRows = sheetAt.getPhysicalNumberOfRows();
//  for (int i = 0; i < numberOfRows; i++) {// 0123
		Row row = sheetAt.getRow(2);
		int numberOfCells = row.getPhysicalNumberOfCells();
		for (int j = 0; j < numberOfCells; j++) {
			Cell cell = row.getCell(j);
			CellType cellType = cell.getCellType();
			if (cellType.equals(CellType.STRING)) {
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
			}
			else if (cellType.equals(CellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				int value=(int) numericCellValue;
				System.out.println(value);
				
			}
		}
	}
public static void particular_column() throws Throwable {
	System.out.println();
	System.out.println("particular column data");
File f =new File("C:\\Users\\sridhar\\eclipse-workspace\\DataDriven-Framework\\Datadrivenread.xlsx");
FileInputStream fis=new FileInputStream(f);
Workbook w=new XSSFWorkbook(fis);
Sheet sheetAt = w.getSheetAt(0);
int numberOfRows = sheetAt.getPhysicalNumberOfRows();
for (int i = 0; i <numberOfRows; i++) {//0123
	Row row = sheetAt.getRow(i);
	int numberOfCells = row.getPhysicalNumberOfCells();
	for (int j = 1; j <numberOfCells; j++) {//0,1
		
		//if (j==2) {
			//break;
		//}
		Cell cell = row.getCell(1);
		CellType cellType = cell.getCellType();
		if (cellType.equals(CellType.STRING)) {
		String stringCellValue = cell.getStringCellValue();	
		System.out.println(stringCellValue);
		}
		else if (cellType.equals(CellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			int value=(int) numericCellValue;
			System.out.println(value);
		}
		}
	}
}




	
//}
	public static void main(String[] args) throws Throwable {
particular_Data();
allData();
particular_Row();
particular_column();
	}

}
