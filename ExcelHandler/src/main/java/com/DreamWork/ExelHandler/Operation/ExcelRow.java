package com.DreamWork.ExelHandler.Operation;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;

public class ExcelRow {

	ExcelRow cRef = null;
	Workbook workbook;
	Sheet sheet;
	String fileName = "";
	Row row;
	int rowNumber;

	ExcelRow() {
		System.out.println("C");
		this.cRef = this;
	}

	ExcelRow(String fileName, Workbook workbook, Sheet sheet) {
		this.workbook = workbook;
		this.fileName = fileName;
		this.sheet = sheet;
		this.cRef = this;
	}

	public ExcelRow excelRow() {
		System.out.println("C function");
		return cRef;
	}

	public ExcelRow createRow(int rowNumber) throws IOException {
		System.out.println("Row created");
		Row row = this.sheet.createRow(rowNumber);
		OutputStream fileOut = new FileOutputStream(this.fileName);
		this.workbook.write(fileOut);
		this.row = row;
		this.rowNumber = rowNumber;
		return cRef;
	}

	public ExcelRow getRow(int rowIndex) throws FileNotFoundException, IOException {
		// remove values from third row but keep third row blank
		Row row = this.sheet.getRow(rowIndex);
		if (row != null)
			this.row = row;
		return cRef;
	}

	public ExcelRow removeRow(int rowIndex) throws FileNotFoundException, IOException {

		// total no. of rows
		int totalRows = this.sheet.getLastRowNum();
		System.out.println("Total no of rows : " + totalRows);

		// remove values from third row but keep third row blank
		if (this.sheet.getRow(rowIndex) != null)
			this.sheet.removeRow(this.sheet.getRow(rowIndex));

		// remove third row completely - 2 for third row and +1; 2+1=3
		// sheet.shiftRows(3, totalRows, -1);
		try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
			this.workbook.write(fileOut);
		}
		return cRef;
	}

	private ExcelRow deleteColumn(int columnToDelete) throws FileNotFoundException, IOException {
		for (int rId = 0; rId < sheet.getLastRowNum(); rId++) {
			Row row = this.sheet.getRow(rId);
			for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
				Cell cOld = row.getCell(cID);
				if (cOld != null) {
					row.removeCell(cOld);
				}
				Cell cNext = row.getCell(cID + 1);
				if (cNext != null) {
					Cell cNew = row.createCell(cID, cNext.getCellType());
					cloneCell(cNew, cNext);
					// Set the column width only on the first row.
					// Other wise the second row will overwrite the original column width set
					// previously.
					if (rId == 0) {
						sheet.setColumnWidth(cID, this.sheet.getColumnWidth(cID + 1));

					}
				}
			}
		}
		try (FileOutputStream fileOut = new FileOutputStream(this.fileName)) {
			this.workbook.write(fileOut);
		}
		return cRef;
	}

	private void cloneCell(Cell cNew, Cell cOld) {
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());

		if (Cell.CELL_TYPE_BOOLEAN == cNew.getCellType()) {
			cNew.setCellValue(cOld.getBooleanCellValue());
		} else if (Cell.CELL_TYPE_NUMERIC == cNew.getCellType()) {
			cNew.setCellValue(cOld.getNumericCellValue());
		} else if (Cell.CELL_TYPE_STRING == cNew.getCellType()) {
			cNew.setCellValue(cOld.getStringCellValue());
		} else if (Cell.CELL_TYPE_ERROR == cNew.getCellType()) {
			cNew.setCellValue(cOld.getErrorCellValue());
		} else if (Cell.CELL_TYPE_FORMULA == cNew.getCellType()) {
			cNew.setCellValue(cOld.getCellFormula());
		}
	}

	public ExcelRow readRow(int rowIndex) throws FileNotFoundException, IOException {

		int totalRows = this.sheet.getLastRowNum();
		System.out.println("Total no of rows : " + totalRows);

		Row row = this.sheet.getRow(rowIndex);
		// remove values from third row but keep third row blank
		if (row != null) {
			Map<Integer, Object> map = new LinkedHashMap<>();
			int lastColumnIndex = row.getLastCellNum();
			for (int i = 0; i < lastColumnIndex; i++) {
				Cell cell = row.getCell(i);
				String type = getCellType(cell.getCellType());
				if (type.equalsIgnoreCase("Numeric") || type.equalsIgnoreCase("Double")) {
					map.put(i, cell.getNumericCellValue());
				} else if (type.equalsIgnoreCase("Boolean")) {
					map.put(i, cell.getBooleanCellValue());
				} else if (type.equalsIgnoreCase("String")) {
					if (isValidDate(cell.getStringCellValue()))
						map.put(i, cell.getDateCellValue());
					map.put(i, cell.getStringCellValue());
				}
			}

			for (Map.Entry<Integer, Object> m2 : map.entrySet()) {
				System.out.println(
						"Key " + m2.getKey() + " Value " + m2.getValue() + " type " + m2.getValue().getClass());
			}
		}

		return cRef;
	}

	public static boolean isValidDate(String inDate) {
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss:ms");
		dateFormat.setLenient(false);
		try {
			dateFormat.parse(inDate.trim());
		} catch (ParseException pe) {
			return false;
		}
		return true;
	}

	private String getCellType(int cellType) {

		switch (cellType) {
		case Cell.CELL_TYPE_NUMERIC:
			return "Numeric";
		case Cell.CELL_TYPE_BOOLEAN:
			return "Boolean";
		case Cell.CELL_TYPE_STRING:
			return "String";
		case Cell.CELL_TYPE_ERROR:
			return "Error";
		default:
			return "String";
		}
	}

	public void writeColumn(int colIndex, Map<Integer, Object> map) throws FileNotFoundException, IOException {
		for (Map.Entry<Integer, Object> m2 : map.entrySet()) {
			System.out.println("Key " + m2.getKey() + " Value " + m2.getValue());
		}
		// remove values from third row but keep third row blank
		for (Map.Entry<Integer, Object> m : map.entrySet()) {
			Row row = this.sheet.createRow(m.getKey());
			Cell cell = row.createCell(colIndex);
			setCellTypeAndValue(cell, m.getValue());
		}
		try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
			this.workbook.write(fileOut);
		}
	}

	public void writeRow(int rowIndex, Map<Integer, Object> map) throws FileNotFoundException, IOException {
		for (Map.Entry<Integer, Object> m2 : map.entrySet()) {
			System.out.println("Key " + m2.getKey() + " Value " + m2.getValue());
		}
		Row row = this.sheet.getRow(rowIndex);
		if (row == null)
			row = this.sheet.createRow(rowIndex);
		// remove values from third row but keep third row blank
		if (row != null) {
			int lastColumnIndex = map.size();
			for (int i = 0; i < lastColumnIndex; i++) {
				Cell cell = row.createCell(i);
				setCellTypeAndValue(cell, map.get(i));
			}
			try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
				this.workbook.write(fileOut);
			}

		}

	}

	private Cell setCellTypeAndValue(Cell cell, Object value) {
		if (value instanceof Integer) {
			cell.setCellValue(new Integer((int) value));
			return cell;
		}
		if (value instanceof Boolean) {
			cell.setCellValue(new Boolean((boolean) value));
			return cell;
		}
		if (value instanceof Date) {
			try {
				SimpleDateFormat simpleDateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzzz yyyy"); // Existing
																											// //
																											// Pattern
				Date currentdate = simpleDateFormat.parse((String) value.toString()); // Returns Date Format,
				SimpleDateFormat simpleDateFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); // New Pattern
				cell.setCellValue(simpleDateFormat1.format(currentdate));
			} catch (ParseException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return cell;
		}
		if (value instanceof String) {
			cell.setCellValue(new String((String) value));
			return cell;
		}
		if (value instanceof Double) {
			cell.setCellValue(new Double((double) value));
			return cell;
		}
		return cell;
	}

	public ExcelCell then() {
		return new ExcelCell(fileName, workbook, sheet, row);
	}
}
