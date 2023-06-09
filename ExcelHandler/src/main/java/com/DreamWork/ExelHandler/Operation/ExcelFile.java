package com.DreamWork.ExelHandler.Operation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelFile {

	String fileName = "";
	private Workbook workbook;

	ExcelFile aRef = null;

	public ExcelFile() {
		this.aRef = this;
	}

	public ExcelFile createFile(String fileName) throws IOException {
		this.fileName = fileName;
		Workbook workbook = new HSSFWorkbook();
		OutputStream fileOut = new FileOutputStream(fileName);
		workbook.write(fileOut);
		this.workbook = workbook;
		return aRef;
	}

	public ExcelFile getFile(String fileName) throws IOException {
		this.fileName = fileName;
		FileInputStream file = new FileInputStream(new File(fileName));
		Workbook workbook = new HSSFWorkbook(file);
		System.out.println("opened File");
		this.workbook = workbook;
		return aRef;
	}

	public ExcelSheet then() {
		aRef = null;
		return new ExcelSheet(fileName, workbook);
	}
}
