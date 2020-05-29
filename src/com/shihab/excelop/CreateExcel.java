package com.shihab.excelop;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
	
	XSSFWorkbook workbook = new XSSFWorkbook();//created blank workbook
	XSSFSheet sheet = workbook.createSheet("sheet1");//sheet is created
	static String title,author;
	static int price;
	public void createexcel(StringBuilder sb,String filename) throws IOException,FileNotFoundException, FormatError {	
		//style for column header
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();  
        font.setFontHeightInPoints((short)14);   
        font.setBold(true);
        //font.setColor(XSSFColor.toXSSFColor("red"));
        style.setFont(font);
        
		//column header with style
        // 0th row contain title,author,price
		XSSFRow row0 = sheet.createRow(0);
		XSSFCell cell0 = row0.createCell(0);
		cell0.setCellStyle(style);
		cell0.setCellValue("title");
		XSSFCell cell1 = row0.createCell(1);
		cell1.setCellStyle(style);
		cell1.setCellValue("author");
		XSSFCell cell2 = row0.createCell(2);
		cell2.setCellStyle(style);
		cell2.setCellValue("price");
		
		//Converting StringBuilder to String
		String con =sb.toString();
		
		//string is splited with new line \n
		String[] arr = con.split("\n"); 
		
		String[] x;
		//spliting each line with "-" and insert in to required cell
		for (int i = 1; i < arr.length+1; i++) {
			if(arr[i-1].matches("[a-z A-Z0-9]+[-]{1}[a-z A-Z]+[-]{1}[0-9]*[.]?[0-9]*")) {
				x=arr[i-1].split("-");
				title=x[0];
				author=x[1];
				price=Integer.parseInt(x[2]);
				//cell 1
				XSSFRow row1 = sheet.createRow(i);
				XSSFCell cell01 = row1.createCell(0);  
				cell01.setCellValue(title);
				//cell 2
				XSSFCell cell12 = row1.createCell(1);
				cell12.setCellValue(author);
				//cell 3
				XSSFCell cell23 = row1.createCell(2);
				cell23.setCellValue(price);
			}
			else {
				throw new FormatError("FORMAT IS INCORRECT");
			}
		}
		//FileOutputStream object created
		FileOutputStream out = new FileOutputStream(new File(filename));
		workbook.write(out);
		out.close();
		workbook.close();
		
	}
}
