package com.gt.poi;

import static org.junit.Assert.*;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class FileManagerTest {

	private final static String[] testi = { "test.xlsx", "test.xls" };
	private final static String SHEET_NAME = "Sheet1";

	@Test
	public void testAdd() throws Exception {
		ClassLoader classLoader = getClass().getClassLoader();

		for (String nameFile : testi) {
			System.out.println("loading " + nameFile);
			InputStream inputFile = new FileInputStream("src/test/resources/"+nameFile);

			Workbook workbook = FileManager.getWorkbook(inputFile,
					ExcelType.getType(nameFile));

			Sheet sheet = FileManager.getSheet(workbook, SHEET_NAME);

			for (int i = 1; i < 6; i++) {
				assertEquals(FileManager.getCellValue(sheet, "B" + i),i+"");
				System.out.println(FileManager.getCellValue(sheet, "A" + i));
			}

		}

	}
}
