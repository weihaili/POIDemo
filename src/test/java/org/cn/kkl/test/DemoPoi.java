package org.cn.kkl.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class DemoPoi {
	
	/**
	 * @param args
	 * @description
	 * 1. create HSSFWorkbook instance 
	 * 2. create sheet(can set sheet width and height)
	 * 3. create first line(row)
	 * 4. create cell of first line(row) 
	 * 5. write value in first cell of first line(row)
	 * 6. save in file system
	 */
	public static void main(String[] args) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		HSSFSheet sheet = workbook.createSheet("my work sheet");
		sheet.setColumnWidth(0, 5000);
		
		HSSFRow firstRow = sheet.createRow(0);
		
		HSSFCell cell = firstRow.createCell(0);
		
		cell.setCellValue("test");
		
		try {
			workbook.write(new FileOutputStream(new File("D:\\temp1\\poitest.xls")));
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("execute complish");
	}

}
