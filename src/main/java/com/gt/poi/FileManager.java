package com.gt.poi;

import java.io.InputStream;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileManager {

	public static Workbook getWorkbook(InputStream inputStream,
			ExcelType excelType) throws Exception {
		System.out.println("getting workbook type '" + excelType + "' ");

		Workbook workbook = null;

		switch (excelType) {
		case HSSF:
			workbook = new HSSFWorkbook(inputStream);
			break;
		case XSSF:
			workbook = new XSSFWorkbook(inputStream);
			break;
		}

		return workbook;
	}

	public static Sheet getSheet(Workbook workbook, String sheetFileName)
			throws Exception {
		System.out.println("geting sheet named '" + sheetFileName + "' ");
		Sheet sheet = workbook.getSheet(sheetFileName);

		if (sheet == null) {
			throw new Exception("sheet with fileName '" + sheetFileName
					+ "' not found");
		}
		return sheet;
	}

	public static String getCellValue(Sheet sheet, String reference) {

		String result = null;

		CellReference cellReference = new CellReference(reference);
		Row row = sheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());

		switch (cell.getCellType()) {
		case STRING:
			result = cell.getStringCellValue();
			break;
		case NUMERIC:
			DecimalFormat formatter = new DecimalFormat("0");
			result = formatter.format(cell.getNumericCellValue());
			break;
		case BOOLEAN:
			result = Boolean.toString(cell.getBooleanCellValue());
			break;
		default:
			result = cell.getStringCellValue();
			break;
		}
		return result;
	}

}
