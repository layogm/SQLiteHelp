package com.Visdrotech.library.sqlitehelp.main;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import android.content.ContentValues;
import android.content.Context;
import android.database.Cursor;
import android.database.sqlite.SQLiteDatabase;
import android.database.sqlite.SQLiteException;
import android.util.Log;

public class SQLiteHelper {

	private SQLiteDatabase database;
	private SQLiteDatabaseHelper dbHelper;
	private Context context;
	private String dbName;

	/*
	 * Constructor to initialise the databse object from the name.
	 * 
	 * @params context, databaseName
	 * 
	 * @return
	 */
	public SQLiteHelper(Context context, String dbName) {
		dbHelper = new SQLiteDatabaseHelper(context, dbName + ".db", 1);
		database = dbHelper.getWritableDatabase();
		this.context = context;
		this.dbName = dbName;
	}

	/*
	 * Method to close the databse
	 * 
	 * @params
	 * 
	 * @return
	 */
	public void close() {
		database.close();
	}

	/*
	 * Method to upgrade the database to newer version
	 * 
	 * @params newdatabse
	 * 
	 * @return true if upgrade is done, false otherwise
	 */
	public boolean upgradeDatabase(SQLiteDatabase newDb) {
		if (newDb.getVersion() > database.getVersion()) {
			dbHelper.onUpgrade(newDb, database.getVersion(), newDb.getVersion());
			return true;
		}
		return false;
	}

	/*
	 * Method to get the current database version
	 * 
	 * @params
	 * 
	 * @return current database version
	 */
	public int getDatabaseVersion() {
		return database.getVersion();
	}

	/*
	 * Method to get the current database version
	 * 
	 * @params
	 * 
	 * @return current database pagesize in bytes
	 */
	public long getDatabasePageSizeInBytes() {
		return database.getPageSize();
	}

	/*
	 * Method to get the current database path
	 * 
	 * @params
	 * 
	 * @return current database path
	 */
	public String getDatabasePath() {
		return database.getPath();
	}

	/*
	 * Method to create new table in the database, if the database for this name
	 * doesn't exist then this method first create this database and then create
	 * the table.
	 * 
	 * @params context, databaseName, tableName, columnNamesList, columnTypeMap
	 * 
	 * @return
	 */
	public void createTable(String tableName, ArrayList<String> columnNames,
			HashMap<String, String> columnDataTypeMap) {
		database.execSQL(generateCreateTableCommandString(tableName,
				columnNames, columnDataTypeMap));
	}

	/*
	 * Method to retrieve all tables names in the current database
	 * 
	 * @params
	 * 
	 * @return ArrayList containing names of all the tables in the database
	 */
	public ArrayList<String> retrieveAllTableNames() {
		ArrayList<String> tableNamesList = new ArrayList<String>();
		Cursor cursor = database.rawQuery(
				"SELECT name FROM sqlite_master WHERE type='table'", null);
		if (cursor.moveToFirst())
			while (!cursor.isAfterLast()) {
				tableNamesList.add(cursor.getString(cursor
						.getColumnIndex("name")));
				cursor.moveToNext();
			}

		return tableNamesList;
	}

	/*
	 * Method to retrieve all tables names in the current database
	 * 
	 * @params tableName
	 * 
	 * @return true if tableName is Present in the database, false otherwise
	 */
	public boolean isTablePresent(String tableName) {
		ArrayList<String> tableNamesList = retrieveAllTableNames();
		return (tableNamesList.contains(tableName));
	}

	/*
	 * Method to add a new column in the existing table in the database
	 * 
	 * @params context, databseName, listOfValuesTobeInserted
	 * 
	 * @return
	 */
	public void addColumn(String tableName, String columnName, String columnType) {

		database.execSQL("ALTER TABLE " + tableName + " ADD " + columnName
				+ " " + columnType + ";");

	}

	/*
	 * Method to change the datatype of an existing column in the existing table
	 * in the database
	 * 
	 * @params context, databseName, listOfValuesTobeInserted
	 * 
	 * @return
	 */
	public void changeColumnDataType(String tableName, String columnName,
			String newColumnType) {

		database.execSQL("ALTER TABLE " + tableName + " ALTER COLUMN "
				+ columnName + " " + newColumnType + ";");

	}

	/*
	 * Method to delete a column from the existing table in the database
	 * 
	 * @params context, databseName, listOfValuesTobeInserted
	 * 
	 * @return
	 */
	public void deleteColumn(String tableName, String columnName) {

		database.execSQL("ALTER TABLE " + tableName + " DROP " + columnName
				+ ";");

	}

	/*
	 * Method to retrieve all column names in the existing table in the database
	 * 
	 * @params context, databseName, tableName
	 * 
	 * @return Array containing names of all the columns of the table
	 */
	public String[] retrieveAllColumnNames(String tableName) {

		Cursor cursor = database.query(tableName, null, null, null, null, null,
				null);
		cursor.moveToFirst();
		String columnNames[] = cursor.getColumnNames();

		return columnNames;
	}

	/*
	 * Method to check if a column by the given column name is present in the
	 * existing table in the database
	 * 
	 * @params context, databseName, tableName, columnName
	 * 
	 * @return true if column with such name exist, false otherwise
	 */
	public boolean isColumnPresent(String tableName, String columnName) {

		String columnsToRetrieve[] = { columnName };
		Cursor cursor = database.query(tableName, columnsToRetrieve, null,
				null, null, null, null);
		return (!cursor.equals(null));
	}

	/*
	 * Method to insert a single value into the table
	 * 
	 * @params context, databseName, tableName, valuesTobeInserted
	 * 
	 * @return
	 */
	public void insertSingleEntry(String tableName, ContentValues value) {

		database.insert(tableName, null, value);

	}

	/*
	 * Method to insert multiple values into the table
	 * 
	 * @params context, databseName,tableName, listOfValuesTobeInserted
	 * 
	 * @return
	 */
	public void insertMultipleEntries(String tableName,
			ArrayList<ContentValues> valueList) {

		for (ContentValues value : valueList)
			database.insert(tableName, null, value);

	}

	/*
	 * Method to update a single value into the table
	 * 
	 * @params context, databseName, tableName, valueTobeUpdated, whereClause
	 * 
	 * @return
	 */
	public void updateSingleEntry(String tableName, ContentValues value,
			String whereClause) {

		database.update(tableName, value, whereClause, null);

	}

	/*
	 * Method to update multiple values into the table
	 * 
	 * @params context, databseName, tableName, listOfValuesTobeUpdated,
	 * listOfWhereClause
	 * 
	 * @return
	 */
	public void updateMultipleEntries(String tableName,
			ArrayList<ContentValues> valueList,
			ArrayList<String> whereClauseList) {

		for (int i = 0; i < valueList.size(); i++)
			database.update(tableName, valueList.get(i),
					whereClauseList.get(i), null);

	}

	/*
	 * Method to delete a single value into the table
	 * 
	 * @params context, databseName, tableName, whereClause
	 * 
	 * @return
	 */
	public void deleteSingleEntry(String tableName, String whereClause) {

		database.delete(tableName, whereClause, null);

	}

	/*
	 * Method to delete multiple values into the table
	 * 
	 * @params context, databseName, tableName, listOfWhereClause
	 * 
	 * @return
	 */
	public void deleteMultipleEntries(String tableName,
			ArrayList<String> whereClauseList) {

		for (String whereClause : whereClauseList)
			database.delete(tableName, whereClause, null);

	}

	/*
	 * Method to delete all values into the table
	 * 
	 * @params context, databseName, tableName
	 * 
	 * @return
	 */
	public void deleteAllEntries(String tableName) {

		database.delete(tableName, null, null);

	}

	/*
	 * Method to retrieve a single value into the table
	 * 
	 * @params context, databseName, tableName, whereClause,
	 * ArrayOfColumnsToRetrieve
	 * 
	 * @return Value of the columns mapped.
	 */
	public ContentValues retrieveSingleEntry(String tableName,
			String whereClause, String[] columnsToRetrieve)
			throws NullPointerException, SQLiteException {

		Cursor cursor = database.query(tableName, columnsToRetrieve,
				whereClause, null, null, null, null);
		ContentValues value = new ContentValues();
		if (!cursor.equals(null) && cursor.getCount() > 0) {
			cursor.moveToFirst();
			for (int i = 0; i < cursor.getColumnCount(); i++)
				value.put(cursor.getColumnName(i), cursor.getString(i));

		} else {
			value = null;
		}

		return value;
	}

	/*
	 * Method to retrieve multiple values into the table
	 * 
	 * @params context, databseName, tableName,listOfWhereClause
	 * ArrayOfcolumnsToRetrieve
	 * 
	 * @return List of values of the columns mapped
	 */
	public ArrayList<ContentValues> retrieveMultipleEntries(String tableName,
			ArrayList<String> whereClauseList, String[] columnsToRetrieve) {

		ArrayList<ContentValues> valueList = new ArrayList<ContentValues>();
		for (String whereClause : whereClauseList) {
			Cursor cursor = database.query(tableName, columnsToRetrieve,
					whereClause, null, null, null, null);
			ContentValues value = new ContentValues();
			if (!cursor.equals(null) && cursor.getCount() > 0) {
				cursor.moveToFirst();
				for (int i = 0; i < columnsToRetrieve.length; i++)
					value.put(columnsToRetrieve[i], cursor.getString(cursor
							.getColumnIndex(columnsToRetrieve[i])));
				valueList.add(value);
			} else {
				value = null;
			}

		}
		return valueList;
	}

	/*
	 * Method to retrieve all values into the table
	 * 
	 * @params context, databseName, tableName, ArrayOfcolumnsToRetrieve
	 * 
	 * @return List of values of the columns mapped
	 */
	public ArrayList<ContentValues> retrieveAllEntries(String tableName,
			String[] columnsToRetrieve) {

		ArrayList<ContentValues> valueList = new ArrayList<ContentValues>();
		Cursor cursor = database.query(tableName, columnsToRetrieve, null,
				null, null, null, null);
		for (cursor.moveToFirst(); !cursor.isAfterLast(); cursor.moveToNext()) {
			{
				ContentValues value = new ContentValues();
				if (!cursor.equals(null) && cursor.getCount() > 0) {
					cursor.moveToFirst();
					for (int i = 0; i < columnsToRetrieve.length; i++)
						value.put(columnsToRetrieve[i], cursor.getString(cursor
								.getColumnIndex(columnsToRetrieve[i])));
					valueList.add(value);
				} else {
					value = null;
				}
			}

		}
		return valueList;
	}

	/*
	 * Method to retrieve all values into the table
	 * 
	 * @params command to execute
	 * 
	 * @return
	 */
	public void executeSqlCommand(String command) {
		database.execSQL(command);
	}

	/*
	 * Method to retrieve all values into the table
	 * 
	 * @params command to execute
	 * 
	 * @return cursor object
	 */
	public Cursor executeSqlCommandWithCursorReturn(String command) {
		return database.rawQuery(command, null);
	}

	/*
	 * Method to generate the SQL command to be executed to create a new table
	 * 
	 * @params tableName, columnNamesList, columnTypeMap
	 * 
	 * @return string to be executed to create a table
	 */
	private String generateCreateTableCommandString(String tableName,
			ArrayList<String> columnNames,
			HashMap<String, String> columnDataTypeMap) {
		StringBuilder sb = new StringBuilder();
		sb.append("CREATE TABLE IF NOT EXISTS " + tableName + " (");
		for (int i = 0; i < columnNames.size(); i++) {
			sb.append(columnNames.get(i) + " "
					+ columnDataTypeMap.get(columnNames.get(i)));
			if (i < (columnNames.size() - 1))
				sb.append(", ");
		}
		sb.append("); ");
		return sb.toString();

	}

	/*
	 * Method to save a table of the database as an Excel file at the specified
	 * filepath
	 * 
	 * @params tableName, filePath, fileName
	 * 
	 * @return
	 */
	public void saveSingleTableAsExcelFile(String tableName, String filePath,
			String fileName) throws RowsExceededException, WriteException,
			IOException {

		ExcelWriteHandler.saveSingleTableData(fileName, filePath, tableName,
				this);
	}

	/*
	 * Method to save the complete database as a single Excel file at the
	 * specified filepath with separate worksheets representing different tables
	 * 
	 * @params filePath, fileName, SqliteHelper Object
	 * 
	 * @return
	 */
	public void saveCompleteDatabaseAsSingleExcelFile(String filePath,
			String fileName) throws RowsExceededException, WriteException,
			IOException {
		Log.d("Log", "A " + getDatabasePageSizeInBytes());

		ExcelWriteHandler.saveAllTablesDataSingleFile(fileName, filePath, this);
	}

	/*
	 * Method to save the complete database as a separate Excel files
	 * representing different tables saved as the <TABLENAME>.xls at specified
	 * location
	 * 
	 * @params filePath, SqliteHelper Object
	 * 
	 * @return
	 */
	public void saveCompleteDatabaseAsMultipleExcelFiles(String tableName,
			String filePath, String fileName) throws RowsExceededException,
			WriteException, IOException {
		ExcelWriteHandler.saveAllTablesDataSeparateFiles(filePath, this);
	}

	/*
	 * Method to add a table from the sheet of the Excel file.
	 * 
	 * @params filePath, fileName, sheetName
	 * 
	 * @return
	 */
	public void addSingleTableFromExcelFile(String sheetName, String filePath,
			String fileName) throws RowsExceededException, WriteException,
			IOException, BiffException {
		ExcelReadHandler.addSingleTableFromExcelSheet(fileName, filePath,
				sheetName, this);
	}

	/*
	 * Method to add multiple tables from multiple sheets in a Excel file.
	 * 
	 * @params filePath, fileName, sheetNames
	 * 
	 * @return
	 */
	public void addMultipleTablesFromExcelFile(ArrayList<String> sheetNames,
			String filePath, String fileName) throws RowsExceededException,
			WriteException, IOException, BiffException {
		ExcelReadHandler.addMultipleTablesFromExcelSheet(fileName, filePath,
				sheetNames, this);
	}

	/*
	 * Method to add multiple tables from all the sheets in a Excel file.
	 * 
	 * @params filePath, fileName
	 * 
	 * @return
	 */
	public void addAllTableFromExcelFile(String sheetName, String filePath,
			String fileName) throws RowsExceededException, WriteException,
			IOException, BiffException {
		ExcelReadHandler.addAllTableFromExcelSheet(fileName, filePath, this);
	}

}
