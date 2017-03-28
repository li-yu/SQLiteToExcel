package com.liyu.sqlitetoexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import android.content.Context;
import android.database.Cursor;
import android.database.sqlite.SQLiteDatabase;
import android.os.Environment;
import android.os.Handler;
import android.os.Looper;

/**
 * Created by liyu on 2015-9-8
 */
public class SqliteToExcel {

    private static Handler handler = new Handler(Looper.getMainLooper());

    private Context mContext;
    private SQLiteDatabase database;
    private String mDbName;
    private ExportListener mListener;
    private String mExportPath;
    private HSSFWorkbook workbook;

    public SqliteToExcel(Context context, String dbName) {
        this(context, dbName, Environment.getExternalStorageDirectory().toString() + File.separator);
    }

    public SqliteToExcel(Context context, String dbName, String exportPath) {
        mContext = context;
        mDbName = dbName;
        mExportPath = exportPath;
        try {
            database = SQLiteDatabase.openOrCreateDatabase(mContext.getDatabasePath(mDbName).getAbsolutePath(), null);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private ArrayList<String> getAllTables() {
        ArrayList<String> tables = new ArrayList<>();
        Cursor cursor = database.rawQuery("select name from sqlite_master where type='table' order by name", null);
        while (cursor.moveToNext()) {
            tables.add(cursor.getString(0));
        }
        cursor.close();
        return tables;
    }

    private ArrayList<String> getColumns(String table) {
        ArrayList<String> columns = new ArrayList<>();
        Cursor cursor = database.rawQuery("PRAGMA table_info(" + table + ")", null);
        while (cursor.moveToNext()) {
            columns.add(cursor.getString(1));
        }
        cursor.close();
        return columns;
    }

    private void exportTables(List<String> tables, final String fileName) throws Exception {
        workbook = new HSSFWorkbook();
        for (int i = 0; i < tables.size(); i++) {
            HSSFSheet sheet = workbook.createSheet(tables.get(i));
            createSheet(tables.get(i), sheet);
        }
        File file = new File(mExportPath, fileName);
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        if (fos != null) {
            fos.flush();
            fos.close();
        }
        workbook.close();
        database.close();
        if (mListener != null) {
            handler.post(new Runnable() {
                @Override
                public void run() {
                    mListener.onCompleted(mExportPath + fileName);
                }
            });
        }
    }

    public void startExportSingleTable(final String table, final String fileName, ExportListener listener) {
        List<String> tables = new ArrayList<>();
        tables.add(table);
        startExportTables(tables, fileName, listener);
    }

    public void startExportAllTables(final String fileName, ExportListener listener) {
        ArrayList<String> tables = getAllTables();
        startExportTables(tables, fileName, listener);
    }

    public void startExportTables(final List<String> tables, final String fileName, ExportListener listener) {
        mListener = listener;
        mListener.onStart();
        new Thread(new Runnable() {

            @Override
            public void run() {
                try {
                    exportTables(tables, fileName);
                } catch (final Exception e) {
                    if (database != null && database.isOpen()) {
                        database.close();
                    }
                    if (mListener != null)
                        handler.post(new Runnable() {
                            @Override
                            public void run() {
                                mListener.onError(e);
                            }
                        });
                }
            }
        }).start();
    }

    private void createSheet(String table, HSSFSheet sheet) {
        HSSFRow rowA = sheet.createRow(0);
        ArrayList<String> columns = getColumns(table);
        for (int i = 0; i < columns.size(); i++) {
            HSSFCell cellA = rowA.createCell(i);
            cellA.setCellValue(new HSSFRichTextString("" + columns.get(i)));
        }
        insertItemToSheet(table, sheet, columns);
    }

    private void insertItemToSheet(String table, HSSFSheet sheet, ArrayList<String> columns) {
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        Cursor cursor = database.rawQuery("select * from " + table, null);
        cursor.moveToFirst();
        int n = 1;
        while (!cursor.isAfterLast()) {
            HSSFRow rowA = sheet.createRow(n);
            for (int j = 0; j < columns.size(); j++) {
                HSSFCell cellA = rowA.createCell(j);
                if (cursor.getType(j) == Cursor.FIELD_TYPE_BLOB) {
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) j, n, (short) (j + 1), n + 1);
                    anchor.setAnchorType(3);
                    patriarch.createPicture(anchor, workbook.addPicture(cursor.getBlob(j), HSSFWorkbook.PICTURE_TYPE_JPEG));
                } else {
                    cellA.setCellValue(new HSSFRichTextString(cursor.getString(j)));
                }
            }
            n++;
            cursor.moveToNext();
        }
        cursor.close();
    }

    public interface ExportListener {
        void onStart();

        void onCompleted(String filePath);

        void onError(Exception e);
    }

}
