package com.Visdrotech.library.sqlitehelp.main;

import android.content.Context;
import android.database.sqlite.SQLiteDatabase;
import android.database.sqlite.SQLiteOpenHelper;
import android.util.Log;

class SQLiteDatabaseHelper extends SQLiteOpenHelper{
	
	public SQLiteDatabaseHelper(Context context, String databaseName,int databaseVersion ) {
		super(context, databaseName, null, databaseVersion );
	}

	@Override
	public void onCreate(SQLiteDatabase db) {
	}

	@Override
	public void onUpgrade(SQLiteDatabase db, int oldVersion, int newVersion) {
		db.execSQL("DROP TABLE IF EXISTS ");
		onCreate(db);
	}

}
