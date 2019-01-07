package com.gt.poi;

import org.apache.poi.ss.formula.eval.NotImplementedException;

public enum ExcelType {
	XSSF, HSSF;

	public static ExcelType getType(String nameFile) {
		if (nameFile.trim().toLowerCase().endsWith("xlsx")) {
			return XSSF;
		} else if (nameFile.trim().toLowerCase().endsWith("xls")) {
			return HSSF;
		}
		throw new NotImplementedException(nameFile);
	}
}
