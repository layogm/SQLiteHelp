package com.Visdrotech.library.sqlitehelp.main;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import android.content.ContentValues;
import android.util.Log;
import android.widget.Toast;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelReadHandler {

	private static ArrayList<ContentValues> getDataFromSheet(Sheet sheet)
			throws RowsExceededException, WriteException {
		ArrayList<ContentValues> valuesList = new ArrayList<ContentValues>();
		ArrayList<String> columnNames = new ArrayList<String>();
		int numCols = sheet.getColumns();
		int numRows = sheet.getRows();
		for (int i = 0; i < numCols; i++) {
			columnNames.add(sheet.getCell(0, i).getContents());
		}
		for (int i = 1; i < numRows; i++) {
			ContentValues value = new ContentValues();
			for (int j = 0; i < numCols; j++) {
				value.put(columnNames.get(j), sheet.getCell(i, j).getContents());
			}
			valuesList.add(value);
		}
		return valuesList;
	}

	private static ArrayList<String> getAllColumnsFromSheet(Sheet sheet)
			throws RowsExceededException, WriteException {
		int numCols = sheet.getColumns();
		ArrayList<String> columnNames = new ArrayList<String>();
		for (int i = 0; i < numCols; i++) {
			columnNames.add(sheet.getCell(0, i).getContents());
		}
		return columnNames;
	}

	public static void fillCompleteDatabaseFromExcelFile(SQLiteHelper helper)
			throws RowsExceededException, WriteException {

	}

	public static void addSingleTableFromExcelSheet(String fileName,
			String filePath, String sheetName, SQLiteHelper helper)
			throws IOException, BiffException, WriteException {
		Workbook workbook = Workbook.getWorkbook(new File(filePath));
		Sheet sheet = workbook.getSheet(sheetName);
		ArrayList<String> columnNames = getAllColumnsFromSheet(sheet);
		HashMap<String, String> columnDataTypeMap = new HashMap<String, String>();
		helper.createTable(sheet.getName(), columnNames, columnDataTypeMap);

		ArrayList<ContentValues> values = new ArrayList<ContentValues>();
		int numRows = sheet.getRows();
		for (int i = 1; i < numRows; i++) {
			ContentValues v = new ContentValues();
			int numCols = sheet.getColumns();
			for (int j = 0; j < numCols; j++) {
				v.put(columnNames.get(j), sheet.getCell(i, j).getContents());
			}
			values.add(v);
		}
		helper.insertMultipleEntries(sheet.getName(), values);
	}

	public static void addAllTableFromExcelSheet(String fileName,
			String filePath, SQLiteHelper helper) throws IOException,
			BiffException, WriteException {
		Workbook workbook = Workbook.getWorkbook(new File(filePath));
		Sheet[] sheets = workbook.getSheets();

		for (Sheet sheet : sheets) {
			ArrayList<String> columnNames = getAllColumnsFromSheet(sheet);
			HashMap<String, String> columnDataTypeMap = new HashMap<String, String>();
			helper.createTable(sheet.getName(), columnNames, columnDataTypeMap);

			ArrayList<ContentValues> values = new ArrayList<ContentValues>();
			int numRows = sheet.getRows();
			for (int i = 1; i < numRows; i++) {
				ContentValues v = new ContentValues();
				int numCols = sheet.getColumns();
				for (int j = 0; j < numCols; j++) {
					v.put(columnNames.get(j), sheet.getCell(i, j).getContents());
				}
				values.add(v);
			}
			helper.insertMultipleEntries(sheet.getName(), values);
		}
	}

	public static void addMultipleTablesFromExcelSheet(String fileName,
			String filePath, ArrayList<String> sheetNames, SQLiteHelper helper)
			throws IOException, BiffException, WriteException {
		Workbook workbook = Workbook.getWorkbook(new File(filePath));
		for (String sheetName : sheetNames) {
			Sheet sheet = workbook.getSheet(sheetName);
			ArrayList<String> columnNames = getAllColumnsFromSheet(sheet);
			HashMap<String, String> columnDataTypeMap = new HashMap<String, String>();
			helper.createTable(sheet.getName(), columnNames, columnDataTypeMap);

			ArrayList<ContentValues> values = new ArrayList<ContentValues>();
			int numRows = sheet.getRows();
			for (int i = 1; i < numRows; i++) {
				ContentValues v = new ContentValues();
				int numCols = sheet.getColumns();
				for (int j = 0; j < numCols; j++) {
					v.put(columnNames.get(j), sheet.getCell(i, j).getContents());
				}
				values.add(v);
			}
			helper.insertMultipleEntries(sheet.getName(), values);
		}
	}
}
