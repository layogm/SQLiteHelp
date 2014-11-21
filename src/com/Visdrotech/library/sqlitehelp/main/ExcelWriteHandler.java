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
import jxl.write.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelWriteHandler {

	private static WritableSheet addDataToSheet(WritableSheet sheet,
			String[] columnNames, ArrayList<ContentValues> valuesList)
			throws RowsExceededException, WriteException {

		// Creating Writable font to be used in the report
		WritableFont normalFont = new WritableFont(
				WritableFont.createFont("MS Sans Serif"),
				WritableFont.DEFAULT_POINT_SIZE, WritableFont.NO_BOLD, false,
				UnderlineStyle.NO_UNDERLINE);

		// creating plain format to write data in excel sheet
		WritableCellFormat normalFormat = new WritableCellFormat(normalFont);
		normalFormat.setWrap(true);
		normalFormat.setAlignment(jxl.format.Alignment.LEFT);
		normalFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		normalFormat.setBorder(jxl.format.Border.NONE,
				jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK);

		// Creating Writable font to be used in the report
		WritableFont headingFont = new WritableFont(
				WritableFont.createFont("Times New Roman"),
				WritableFont.DEFAULT_POINT_SIZE, WritableFont.BOLD, false,
				UnderlineStyle.NO_UNDERLINE, Colour.BLACK);

		// creating plain format to write data in excel sheet
		WritableCellFormat headingFormat = new WritableCellFormat(headingFont);
		headingFormat.setWrap(true);
		headingFormat.setBackground(Colour.GRAY_25);
		headingFormat.setAlignment(jxl.format.Alignment.LEFT);
		headingFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		headingFormat.setBorder(jxl.format.Border.ALL,
				jxl.format.BorderLineStyle.DOUBLE, jxl.format.Colour.BLACK);

		for (int i = 0; i < columnNames.length; i++) {
			sheet.addCell(new jxl.write.Label(i, 0, columnNames[i],
					headingFormat));
		}

		for (int i = 0; i < valuesList.size(); i++) {
			ContentValues value = valuesList.get(i);

			// Number of columns and the number of values in contentValues
			// should be same at every step
			if (columnNames.length != value.size())
				return null;

			for (int j = 0; j < value.size(); j++) {
				sheet.addCell(new jxl.write.Label(j, i + 1, value
						.getAsString(columnNames[j]), normalFormat));
			}

		}

		return sheet;
	}

	public static void saveSingleTableData(String fileName, String filePath,
			String tableName, SQLiteHelper helper) throws IOException,
			RowsExceededException, WriteException {
		WorkbookSettings ws = new WorkbookSettings();
		WritableWorkbook workbook = Workbook.createWorkbook(new File(filePath,
				fileName), ws);
		WritableSheet sheet = workbook.createSheet(tableName, 0);
		String[] colNames = helper.retrieveAllColumnNames(tableName);
		sheet = addDataToSheet(sheet, colNames,
				helper.retrieveAllEntries(tableName, colNames));
		workbook.write();
		workbook.close();
	}

	public static void saveAllTablesDataSingleFile(String fileName,	String filePath	, SQLiteHelper helper) throws IOException,RowsExceededException, WriteException {
		WorkbookSettings ws = new WorkbookSettings();
		WritableWorkbook workbook = Workbook.createWorkbook(new File(filePath,
				fileName), ws);
		ArrayList<String> tableNamesList = helper.retrieveAllTableNames();
		for (String tableName : tableNamesList) {
			WritableSheet sheet = workbook.createSheet(tableName, 0);
			String[] colNames = helper.retrieveAllColumnNames(tableName);
			sheet = addDataToSheet(sheet, colNames,
					helper.retrieveAllEntries(tableName, colNames));
		}
		workbook.write();
		workbook.close();
	}

	public static void saveAllTablesDataSeparateFiles(String filePath,
			SQLiteHelper helper) throws IOException, RowsExceededException,
			WriteException {
		ArrayList<String> tableNamesList = helper.retrieveAllTableNames();
		for (String tableName : tableNamesList) {
			WorkbookSettings ws = new WorkbookSettings();
			WritableWorkbook workbook = Workbook.createWorkbook(new File(
					filePath, tableName + ".xls"), ws);
			WritableSheet sheet = workbook.createSheet(tableName, 0);
			String[] colNames = helper.retrieveAllColumnNames(tableName);
			sheet = addDataToSheet(sheet, colNames,
					helper.retrieveAllEntries(tableName, colNames));
			workbook.write();
			workbook.close();
		}
	}
}
