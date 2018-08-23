package org.cn.kkl.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author Admin create order template
 */
public class CreateOrderXls {

	public static void main(String[] args) {
		CreateOrderXls xls = new CreateOrderXls();
		xls.createTemplate();

		System.out.println("execute completely");
	}

	public void setValue() {
		HSSFWorkbook book = cellMerge("value");
		HSSFSheet sheet = book.getSheet("value");

		sheet.createRow(0).createCell(0).setCellValue("purchase order");

		sheet.getRow(2).getCell(0).setCellValue("supplier");

		sheet.getRow(3).getCell(0).setCellValue("order date");
		sheet.getRow(3).getCell(2).setCellValue("manager");

		sheet.getRow(4).getCell(0).setCellValue("check date");
		sheet.getRow(4).getCell(2).setCellValue("manager");

		sheet.getRow(5).getCell(0).setCellValue("purchase date");
		sheet.getRow(5).getCell(2).setCellValue("manager");

		sheet.getRow(6).getCell(0).setCellValue("inStore date");
		sheet.getRow(6).getCell(2).setCellValue("manager");

		sheet.getRow(7).getCell(0).setCellValue("order detail");

		sheet.getRow(8).getCell(0).setCellValue("goods name");
		sheet.getRow(8).getCell(0).setCellValue("quantity");
		sheet.getRow(8).getCell(0).setCellValue("price");
		sheet.getRow(8).getCell(0).setCellValue("amount");

	}

	/**
	 * merge cell
	 */
	public HSSFWorkbook cellMerge(String sheetName) {
		//HSSFWorkbook book = createTemplate();
		HSSFSheet sheet = new HSSFWorkbook().createSheet();
		
		//HSSFSheet sheet = book.getSheet(sheetName);
		// the first four columns of the first row are merged
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

		// second columns to fourth columns of third row are merged
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 3));

		// first column to fourth column of eighth row are merged
		sheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 3));

		return null;
	}

	/**
	 * create rows 10 and columns 4 sheet
	 * 
	 * @param sheetName
	 * @return
	 */
	public void createTemplate() {
		HSSFWorkbook book = new HSSFWorkbook();
		HSSFSheet sheet = book.createSheet("valueSet");

		HSSFCellStyle cellStyle = book.createCellStyle();
		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		
		//content vertical center and align center
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		
		//set content font
		HSSFFont font = book.createFont();
		font.setFontName("SimSun");
		font.setFontHeightInPoints((short) 11);
		cellStyle.setFont(font);
		
		//title style
		HSSFCellStyle titleStyle = book.createCellStyle();
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		HSSFFont titleFont = book.createFont();
		titleFont.setFontName("boldface");
		titleFont.setBold(true);
		titleFont.setFontHeightInPoints((short) 18);
		titleStyle.setFont(titleFont);
		
		//set date format
		HSSFCellStyle dateStyle = book.createCellStyle();
		dateStyle.cloneStyleFrom(cellStyle);
		HSSFDataFormat dataFormat = book.createDataFormat();
		dateStyle.setDataFormat(dataFormat.getFormat("yyyy-MM-dd HH:mm:ss"));
		
		
		// the first four columns of the first row are merged
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

		// second columns to fourth columns of third row are merged
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 3));

		// first column to fourth column of eighth row are merged
		sheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 3));
		
		for (int i = 2; i < 12; i++) {
			HSSFRow row = sheet.createRow(i);
			for (int j = 0; j < 4; j++) {
				HSSFCell cell = row.createCell(j);
				cell.setCellStyle(cellStyle);
			}
		}
		
		sheet.createRow(0).createCell(0).setCellValue("purchase order");
		sheet.getRow(0).getCell(0).setCellStyle(titleStyle);
		
		sheet.getRow(2).getCell(0).setCellValue("supplier");

		sheet.getRow(3).getCell(0).setCellValue("order date");
		sheet.getRow(3).getCell(2).setCellValue("manager");

		sheet.getRow(4).getCell(0).setCellValue("check date");
		sheet.getRow(4).getCell(2).setCellValue("manager");

		sheet.getRow(5).getCell(0).setCellValue("purchase date");
		sheet.getRow(5).getCell(2).setCellValue("manager");

		sheet.getRow(6).getCell(0).setCellValue("inStore date");
		sheet.getRow(6).getCell(2).setCellValue("manager");

		sheet.getRow(7).getCell(0).setCellValue("order detail");

		sheet.getRow(8).getCell(0).setCellValue("goods name");
		sheet.getRow(8).getCell(1).setCellValue("quantity");
		sheet.getRow(8).getCell(2).setCellValue("price");
		sheet.getRow(8).getCell(3).setCellValue("amount");
		
		sheet.getRow(0).setHeight((short) 1000);
		for (int i = 2; i < 12; i++) {
			sheet.getRow(i).setHeight((short) 500);
		}
		for (int i = 0; i < 4; i++) {
			sheet.setColumnWidth(i, 7000);
		}
		
		for (int i = 3; i < 7; i++) {
			sheet.getRow(i).getCell(1).setCellStyle(dateStyle);
		}
		sheet.getRow(3).getCell(1).setCellValue(new Date());
		
		writeFileSystem(book);
	}

	/**
	 * @description step: 1. create workBook 2. create sheet 3. create style
	 *              tools instance 4. set workBook up left right down border 5.
	 *              create 10 rows and 4 columns sheet
	 */
	public void createOrderTemplate() {

		HSSFWorkbook book = new HSSFWorkbook();
		HSSFSheet sheet = book.createSheet("purchase order");

		HSSFCellStyle style_content = book.createCellStyle();
		style_content.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style_content.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style_content.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style_content.setBorderRight(HSSFCellStyle.BORDER_THIN);

		// 10 rows 4 columns
		for (int i = 2; i < 12; i++) {
			HSSFRow row = sheet.createRow(i);
			for (int j = 0; j < 4; j++) {
				HSSFCell cell = row.createCell(j);
				cell.setCellStyle(style_content);
			}
		}
		writeFileSystem(book);
	}

	public void writeFileSystem(HSSFWorkbook book) {
		try {
			book.write(new FileOutputStream(new File("D:\\temp1\\template" + Math.random() + ".xls")));
			book.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
