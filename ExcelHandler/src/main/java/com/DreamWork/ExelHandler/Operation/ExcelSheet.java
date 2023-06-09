package com.DreamWork.ExelHandler.Operation;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelSheet {

	ExcelSheet bRef = null;
	String fileName = "";
	Workbook workbook;
	Sheet sheet;

	ExcelSheet() {
		System.out.println("B");
		this.bRef = this;
	}

	ExcelSheet(String fileName, Workbook workbook) {
		this.workbook = workbook;
		this.fileName = fileName;
		this.bRef = this;
	}

	public ExcelSheet createSheet(String sheetName) throws IOException {
		System.out.println("Sheet created");
		Sheet sheet = this.workbook.createSheet(sheetName);
		OutputStream fileOut = new FileOutputStream(this.fileName);
		this.workbook.write(fileOut);
		this.sheet = sheet;
		return bRef;
	}

	public ExcelSheet getSheet(String sheetName) throws IOException {
		System.out.println("get Sheet");
		Sheet sheet = this.workbook.getSheet(sheetName);
		this.sheet = sheet;
		return bRef;
	}

	public ExcelSheet addHeader(String leftHeader, String centerHeader, String rightHeader) {
		try (OutputStream fileOut = new FileOutputStream(this.fileName)) {
			// Creating Header
			Header header = sheet.getHeader();
			if (!leftHeader.isEmpty())
				header.setLeft(leftHeader);
			if (!centerHeader.isEmpty())
				header.setCenter(centerHeader);
			if (!rightHeader.isEmpty())
				header.setRight(rightHeader);
			// Creating Row
			this.workbook.write(fileOut);
			return bRef;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return null;
		}
	}

	public ExcelSheet setPageNumberInExcel() {
		try (OutputStream os = new FileOutputStream(this.fileName)) {
			Footer footer = sheet.getFooter();
			footer.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());
			this.workbook.write(os);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return bRef;
	}

	public ExcelSheet cloneSheet(String sourceSheetName, String destinationSheetName) {
		try (OutputStream os = new FileOutputStream(this.fileName)) {
			int indexOfSourceSheet = workbook.getSheetIndex(sourceSheetName);
			if (indexOfSourceSheet < 0) {
				throw new Exception();
			}
			String des = this.workbook.cloneSheet(indexOfSourceSheet).getSheetName();
			int indexOfNewSheet = workbook.getSheetIndex(des);
			this.workbook.setSheetName(indexOfNewSheet, destinationSheetName);
			this.workbook.write(os);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return bRef;
	}

	public ExcelSheet excelSheet() {
		System.out.println("B function");
		return bRef;
	}

	public ExcelRow then() {
		return new ExcelRow(fileName, workbook, sheet);
	}
}
