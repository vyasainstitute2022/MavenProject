package com.vyasa.maven.MavenProject;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcel {
	
	public static void readExcel(String path,String sname) throws IOException
	{
		FileInputStream inputstream=new FileInputStream(path);
		XSSFWorkbook workbook=new XSSFWorkbook(inputstream);
		XSSFSheet sheet=workbook.getSheet(sname); 
		
		int rowcount=sheet.getLastRowNum(); 
		int colcount=sheet.getRow(0).getLastCellNum(); 		
		for (int r = 0; r < rowcount; r++) {
			for (int c = 0; c < colcount; c++) {
				XSSFCell cell=sheet.getRow(r).getCell(c);   
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
		XSSFWorkbook workbook=new XSSFWorkbook(inputstream);
		XSSFSheet sheet=workbook.getSheet(sname); 
		
		int rowcount=sheet.getLastRowNum();   
		int colcount=sheet.getRow(0).getLastCellNum(); 
		
		XSSFRow newRow=sheet.createRow(rowcount+1);
		for (int j = 0; j < colcount; j++) {
			XSSFCell cell=newRow.createCell(j);
			cell.setCellValue(dataToWtite[j]);   
		}
		inputstream.close();
		FileOutputStream outputstream=new FileOutputStream(path);
		workbook.write(outputstream);
		outputstream.close();
		workbook.close();
	}
	
	public static void main(String[] args) throws Exception {
		String filepath="D:\\JavaProgramming\\MavenProject\\TestData\\TestData.xlsx";
		String readSheetName="ReadData";
		String writeSheetName="WriteData";
		readExcel(filepath,readSheetName);
		writeExcel(filepath,writeSheetName);
	}

}
