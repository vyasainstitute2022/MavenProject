package com.vyasa.maven.MavenProject;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadWriteExcelAll {
	
	public static void readExcel(String path,String sname) throws IOException, EncryptedDocumentException, InvalidFormatException
	{
		FileInputStream inputstream=new FileInputStream(path);
		Workbook workbook=WorkbookFactory.create(inputstream);
		Sheet sheet=workbook.getSheet(sname); 
		
		int rowcount=sheet.getLastRowNum(); 
		int colcount=sheet.getRow(0).getLastCellNum();
		
		for (int r = 0; r <=rowcount; r++) {
			for (int c = 0; c < colcount; c++) {
				Cell cell=sheet.getRow(r).getCell(c);    
				System.out.print(cell+" | ");
			}
			System.out.println();
		}
		workbook.close();
	}
	
	public static void writeExcel(String path,String sname) throws Exception
	{
		String[] dataToWtite={"P","13","8"};   
		FileInputStream inputstream=new FileInputStream(path); 
		Workbook workbook=WorkbookFactory.create(inputstream);
		Sheet sheet=workbook.getSheet(sname); 
		
		int rowcount=sheet.getLastRowNum();   
		int colcount=sheet.getRow(0).getLastCellNum();
		
		Row newRow=sheet.createRow(rowcount+1); 
		for (int j = 0; j < colcount; j++) {
			Cell cell=newRow.createCell(j);
			cell.setCellValue(dataToWtite[j]);   
			}
		inputstream.close();
		FileOutputStream outputstream=new FileOutputStream(path);
		workbook.write(outputstream);
		outputstream.close();
	}
	
	public static void main(String[] args) throws Exception {
		String filepath="D:\\JavaProgramming\\MavenProject\\TestData\\TestDataOld.xls";
		String readSheetName="ReadData";
		String writeSheetName="WriteData";
		readExcel(filepath,readSheetName);
		writeExcel(filepath,writeSheetName);
	}

}
