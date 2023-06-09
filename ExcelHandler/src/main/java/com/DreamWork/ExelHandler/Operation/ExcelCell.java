package com.DreamWork.ExelHandler.Operation;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.comparator.Comparators;


public class ExcelCell {

	ExcelCell dRef = null;
	private Workbook workbook;
	private Sheet sheet;
	private String fileName = "";
	private Row row;
	private int rowNumber;
	private Cell cell;
	private int cellNumber;

	ExcelCell(String fileName, Workbook workbook, Sheet sheet, Row row) {
		this.workbook = workbook;
		this.fileName = fileName;
		this.sheet = sheet;
		this.row = row;
		this.dRef = this;
	}

	public ExcelCell createCell(int cellNumber) throws IOException {
		System.out.println("Cell created");
		Cell cell = this.row.createCell(cellNumber);
		OutputStream fileOut = new FileOutputStream(this.fileName);
		this.workbook.write(fileOut);
		this.cell = cell;
		this.cellNumber = cellNumber;
		return dRef;
	}

	private void insertRowAndColumnWhenEmpty(Sheet sheet, int rowNumber, int colNumber) {
		if (sheet.getRow(rowNumber) == null)
			sheet.createRow(rowNumber);
		if (sheet.getRow(rowNumber).getCell(colNumber) == null)
			sheet.getRow(rowNumber).createCell(colNumber);
	}

	public ExcelCell setCellTypeAndValue(Object value) {
		try (OutputStream os = new FileOutputStream(this.fileName)) {

			Row row = null;
			Cell cell = null;
			if (sheet.getRow(rowNumber) == null)
				row = sheet.createRow(rowNumber);
			else
				row = sheet.getRow(rowNumber);

			if (row.getCell(this.cellNumber) == null)
				cell = row.createCell(this.cellNumber);
			else
				cell = row.getCell(this.cellNumber);

			if (value instanceof Integer)
				cell.setCellValue((Integer) value);
			if (value instanceof Boolean)
				cell.setCellValue((Boolean) value);
			if (value instanceof Date)
				cell.setCellValue((Date) value);
			if (value instanceof String)
				cell.setCellValue((String) value);
			if (value instanceof Double)
				cell.setCellValue((Double) value);
			this.workbook.write(os);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return dRef;
	}

	public ExcelCell insertIntoCell(String value) {
		try (OutputStream os = new FileOutputStream(this.fileName)) {
			insertRowAndColumnWhenEmpty(this.sheet, this.rowNumber, this.cellNumber);
			Cell cell = this.sheet.getRow(this.rowNumber).getCell(this.cellNumber);
			cell.setCellValue(value);
			workbook.write(os);
			System.out.println("insert Successfully");
			return dRef;
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return dRef;
	}

	public void forEach(List<Integer> row, List<Integer> cell, List<String> value) {
		try (OutputStream os = new FileOutputStream(this.fileName)) {

			int valueIndex = 0;
			for (int rIndex = 0; rIndex < row.size(); rIndex++) {
				int temp = rIndex * row.size();
				for (int cIndex = temp; cIndex < temp + row.size(); cIndex++) {
					insertRowAndColumnWhenEmpty(this.sheet, rIndex, cIndex);
					this.sheet.getRow(rIndex).getCell(cIndex).setCellValue(value.get(valueIndex));
					if (value.size() != valueIndex)
						valueIndex++;
				}

			}
			workbook.write(os);
			System.out.println("insert Successfully");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public ExcelCell addBorderToCell() {
		// Workbook wb = new HSSFWorkbook();
		// Sheet sheet = workBook.getSheet(sheetName);
		// Row row = sheet.getRow(rowNumber);
		insertRowAndColumnWhenEmpty(sheet, rowNumber, this.cellNumber);
		Cell cell = sheet.getRow(rowNumber).getCell(cellNumber);
		// Styling border of cell.
		CellStyle style = this.workbook.createCellStyle();
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLUE.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		cell.setCellStyle(style);
		try (OutputStream fileOut = new FileOutputStream(fileName)) {
			this.workbook.write(fileOut);
			return dRef;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return dRef;
		}
	}

	private Short colorIndex(String color) {
		Short colorIndex = null;
		switch (color.toUpperCase()) {
		case "GREEN":
			colorIndex = IndexedColors.GREEN.getIndex();
			break;
		case "YELLOW":
			colorIndex = IndexedColors.YELLOW.getIndex();
			break;
		case "BLUE":
			colorIndex = IndexedColors.BLUE.getIndex();
			break;
		case "WHITE":
			colorIndex = IndexedColors.WHITE.getIndex();
			break;
		case "RED":
			colorIndex = IndexedColors.RED.getIndex();
			break;
		case "ORANGE":
			colorIndex = IndexedColors.ORANGE.getIndex();
			break;
		case "BROWN":
			colorIndex = IndexedColors.BROWN.getIndex();
			break;
		case "GOLD":
			colorIndex = IndexedColors.GOLD.getIndex();
			break;
		}
		return colorIndex;
	}

	public ExcelCell addColorToCell(String foreGroundColor, String backgroudColor) {
		try (OutputStream fileOut = new FileOutputStream(this.fileName)) {
			// Workbook wb = new XSSFWorkbook();
			// Sheet sheet = wb.createSheet("Sheet");
			// Row row = sheet.createRow(1);
			CellStyle style = this.workbook.createCellStyle();
			// Setting Background color
			style.setFillBackgroundColor(colorIndex(backgroudColor));
			style.setFillPattern(CellStyle.BIG_SPOTS);
			Cell cell = sheet.getRow(rowNumber).getCell(this.cellNumber);
			cell.setCellStyle(style);
			// Setting Foreground Color
			style = this.workbook.createCellStyle();
			style.setFillForegroundColor(colorIndex(foreGroundColor));
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(style);
			this.workbook.write(fileOut);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return dRef;
	}

	public ExcelCell mergeCells(int rowStartNumber, int rowEndNumber, int colStartNumber, int colEndNumber) {
		try (OutputStream fileOut = new FileOutputStream(fileName)) {
			this.sheet
					.addMergedRegion(new CellRangeAddress(rowStartNumber, rowEndNumber, colStartNumber, colEndNumber));
			this.workbook.write(fileOut);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return dRef;
	}

	public void end() {

	}

	public ExcelCell wrapText() {
		try (OutputStream fileOut = new FileOutputStream(this.fileName)) {
			Row row = sheet.getRow(this.rowNumber);
			Cell cell = null;
			if (row == null)
				row = sheet.createRow(this.rowNumber);
			cell = row.getCell(this.cellNumber);
			if (cell == null)
				cell = row.createCell(this.cellNumber);

			CellStyle cs = this.workbook.createCellStyle();
			cs.setWrapText(true);
			cell.setCellStyle(cs);
			row.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
			sheet.autoSizeColumn(2);
			this.workbook.write(fileOut);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return dRef;
	}

	public void calCulateSumInSheet(List<Integer> rowNumber, List<Integer> cellNumber, int resultRowNumber,
			int resultCellNumber) {
		try (OutputStream fileOut = new FileOutputStream(this.fileName)) {
			Row row = sheet.getRow(resultRowNumber);
			Cell cell = null;
			if (row == null)
				row = sheet.createRow(resultRowNumber);
			cell = row.getCell(resultCellNumber);
			if (cell == null)
				cell = row.createCell(resultCellNumber);

			rowNumber.sort(Comparator.naturalOrder());
			StringBuilder cellDetail = new StringBuilder("");
			for (int c = 0; c < cellNumber.size(); c++) {
				for (int r = 0; r < rowNumber.size(); r++) {
					String temp = String.valueOf(getColumn(cellNumber.get(c))) + (rowNumber.get(r) + 1) + ",";
					cellDetail.append(temp);
				}
			}
			cell.setCellType(Cell.CELL_TYPE_FORMULA);
			cell.setCellFormula("SUM(" + cellDetail.substring(0, (cellDetail.length() - 1)) + ")");
			this.workbook.write(fileOut);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	private char getColumn(int C) {
		char ch = 0;
		for (int i = 65, temp = 0; i <= 90; i++, temp++) {
			if (C == temp) {
				ch = (char) i;
				break;
			}
		}
		return ch;
	}

	public ExcelCell AddComments(String text, String comment) {
		try (FileOutputStream out = new FileOutputStream(fileName)) {
			HSSFPatriarch hpt = (HSSFPatriarch) sheet.createDrawingPatriarch();
			Row row = sheet.getRow(this.rowNumber);
			Cell cell = null;
			if (row == null)
				row = sheet.createRow(this.rowNumber);
			cell = row.getCell(this.cellNumber);
			if (cell == null)
				cell = row.createCell(this.cellNumber);

			// Setting size and position of the comment in worksheet
			Comment comment1 = hpt.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short) (this.cellNumber + 1),
					(this.rowNumber + 1), (short) (this.cellNumber + 3), this.rowNumber + 5));
			// Setting comment text
			comment1.setString(new HSSFRichTextString(comment));
			// Associating comment to the cell
			cell.setCellComment(comment1);
			this.workbook.write(out);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

		return dRef;
	}

}
