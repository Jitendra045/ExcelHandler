package com.DreamWork.ExcelHandler;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.DreamWork.ExelHandler.Operation.ExcelFile;

public class ExcelMainApplication {

	public static void main(String[] args) throws IOException {
		ExcelFile obj = new ExcelFile();
		String fileName = "test.xls";
		String sheetName = "testSheetfs";
		String sheetName1 = "testSheet1";

		Map<Integer, Object> m = new HashMap<>();
		m.put(0, 800);
		m.put(1, 100);
		m.put(2, 100);
		m.put(3, 100);
		m.put(4, 100);
		m.put(5, 100);
		m.put(6, 100);
		m.put(7, 100);
		m.put(8, 100);
		m.put(9, 100);

		List<Integer> r = new ArrayList<>();
		r.add(0);
		r.add(1);
		r.add(2);
		r.add(3);
		r.add(4);
		r.add(5);
		r.add(6);
		r.add(7);
		r.add(8);
		r.add(9);

		List<Integer> c = new ArrayList<>();
		c.add(1);
//		new ExcelFile().createFile(fileName).then().createSheet(sheetName).then().writeRow(1, m);

//		new ExcelFile().getFile(fileName).then().getSheet(sheetName).then().readRow(1);

		new ExcelFile().createFile(fileName).then().createSheet(sheetName).then().writeColumn(1, m);

//		new ExcelFile().getFile(fileName).then().getSheet(sheetName).then().writeColumn(2, m);
//		new ExcelFile().getFile(fileName).then().getSheet(sheetName).then().writeColumn(3, m);
//		new ExcelFile().getFile(fileName).then().getSheet(sheetName).then().writeColumn(4, m);
		new ExcelFile().getFile(fileName).then().getSheet(sheetName).then().then().calCulateSumInSheet(r, c, 14, 14);
		;
////	;

//		List<String> val = new ArrayList<>();
//		List<Integer> row = new ArrayList<>();
//		List<Integer> cell = new ArrayList<>();
//		for (int r = 0; r < 10; r++) {
//			row.add(r);
//			for (int c = 0; c < 26; c++) {
//				cell.add(c);
//				val.add("Test" + r + c);
//			}
//		}
////		new ExcelFile().createFile(fileName).then().createSheet(sheetName).then().then().forEach(row, cell, val);
////		;
////
////		new ExcelFile().getFile(fileName).then().cloneSheet(sheetName, sheetName1);
	}
}
