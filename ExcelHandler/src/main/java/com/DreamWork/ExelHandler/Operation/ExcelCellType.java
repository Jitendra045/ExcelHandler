package com.DreamWork.ExelHandler.Operation;

import java.util.List;

public class ExcelCellType {

	private String columnName;
	private List<Object> ls;
	private Object type;

	public String getColumnName() {
		return columnName;
	}

	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}

	public List<Object> getLs() {
		return ls;
	}

	public void setLs(List<Object> ls) {
		this.ls = ls;
	}

	public Object getType() {
		return type;
	}

	public void setType(Object type) {
		this.type = type;
	}
}
