package com.liyu.sqlitetoexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import android.content.Context;
import android.database.Cursor;
import android.database.sqlite.SQLiteDatabase;
import android.os.Handler;
import android.os.Looper;
import android.text.TextUtils;

/**
 * SQLite to Excel
 * Created by liyu on 2015-9-8
 */
public class SQLiteToExcel {

    private static Handler handler = new Handler(Looper.getMainLooper());

    private String protectKey;
    private String encryptKey;
    private String fileName;
    private String filePath;
    private List<String> tables;

    private SQLiteDatabase database;
    private Workbook workbook;

    public static class Builder {
        private String dataBaseName;
        private String filePath;
        private String fileName;
        private String protectKey;
        private String encryptKey;
        private List<String> tables;

        public Builder(Context context) {
            this.filePath = context.getExternalFilesDir(null).getPath();
        }

        public SQLiteToExcel build() {
            if (TextUtils.isEmpty(dataBaseName)) {
                throw new IllegalArgumentException("Database name must not be null.");
            }
            if (TextUtils.isEmpty(fileName)) {
                throw new IllegalArgumentException("Output file name must not be null.");
            }
            return new SQLiteToExcel(tables, protectKey, encryptKey, fileName, dataBaseName, filePath);
        }

        public Builder setDataBase(String dataBaseName) {
            this.dataBaseName = dataBaseName;
            this.fileName = new File(dataBaseName).getName() + ".xls";
            return this;
        }

        public Builder setFileName(String fileName) {
            this.fileName = fileName;
            return this;
        }

        public Builder setProtectKey(String protectPassword) {
            this.protectKey = protectPassword;
            return this;
        }

        public Builder setEncryptKey(String encryptKey) {
            this.encryptKey = encryptKey;
            return this;
        }

        public Builder setTables(String... tables) {
            this.tables = Arrays.asList(tables);
            return this;
        }

        public Builder setPath(String path) {
            this.filePath = path;
            return this;
        }

        public String start() {
            final SQLiteToExcel sqliteToExcel = build();
            return sqliteToExcel.start();
        }

        public void start(ExportListener listener) {
            final SQLiteToExcel sqliteToExcel = build();
            sqliteToExcel.start(listener);
        }
    }

    /**
     * importTables task
     *
     * @return output file path
     */
    public String start() {
        if (tables == null || tables.size() == 0) {
            tables = getAllTables(database);
        }
        try {
            return exportTables(tables, fileName);
        } catch (Exception e) {
            if (database != null && database.isOpen()) {
                database.close();
            }
            return null;
        }
    }

    /**
     * importTables task with a listener
     *
     * @param listener
     */
    public void start(final ExportListener listener) {
        if (tables == null || tables.size() == 0) {
            tables = getAllTables(database);
        }
        if (listener != null) {
            listener.onStart();
        }
        new Thread(new Runnable() {

            @Override
            public void run() {
                try {
                    final String filePath = exportTables(tables, fileName);
                    if (listener != null) {
                        handler.post(new Runnable() {
                            @Override
                            public void run() {
                                listener.onCompleted(filePath);
                            }
                        });
                    }
                } catch (final Exception e) {
                    if (database != null && database.isOpen()) {
                        database.close();
                    }
                    if (listener != null)
                        handler.post(new Runnable() {
                            @Override
                            public void run() {
                                listener.onError(e);
                            }
                        });
                }
            }
        }).start();
    }

    private SQLiteToExcel(List<String> tables, String protectKey, String encryptKey, String fileName, String dataBaseName, String filePath) {
        this.protectKey = protectKey;
        this.encryptKey = encryptKey;
        this.fileName = fileName;
        this.filePath = filePath;
        this.tables = tables;

        try {
            database = SQLiteDatabase.openOrCreateDatabase(dataBaseName, null);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * core code, export tables to a excel file
     *
     * @param tables   database tables
     * @param fileName target file name
     * @return target file path
     * @throws Exception
     */
    private String exportTables(List<String> tables, final String fileName) throws Exception {
        if (fileName.toLowerCase().endsWith(".xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("file name is null or unsupported file format!");
        }
        for (int i = 0; i < tables.size(); i++) {
            Sheet sheet = workbook.createSheet(tables.get(i));
            fillSheet(tables.get(i), sheet);
            if (!TextUtils.isEmpty(protectKey)) {
                sheet.protectSheet(protectKey);
            }
        }
        File file = new File(filePath, fileName);
        FileOutputStream fos1 = new FileOutputStream(file);
        workbook.write(fos1);
        if (fos1 != null) {
            fos1.flush();
            fos1.close();
        }
        workbook.close();
        database.close();
        if (!TextUtils.isEmpty(encryptKey)) {
            SecurityUtil.EncryptFile(file, encryptKey);
        }

        return file.getPath();
    }

    /**
     * Query the database ,then fill in to the sheet
     *
     * @param table database table name
     * @param sheet target sheet
     */
    private void fillSheet(String table, Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        ArrayList<String> columns = getTableColumns(database, table);
        for (int i = 0; i < columns.size(); i++) {
            Cell cellA = headerRow.createCell(i);
            cellA.setCellValue(new HSSFRichTextString("" + columns.get(i)));
        }
        Drawing patriarch = sheet.createDrawingPatriarch();
        Cursor cursor = database.rawQuery("select * from " + table, null);
        cursor.moveToFirst();
        int n = 1;
        while (!cursor.isAfterLast()) {
            Row rowA = sheet.createRow(n);
            for (int j = 0; j < columns.size(); j++) {
                Cell cellA = rowA.createCell(j);
                if (cursor.getType(j) == Cursor.FIELD_TYPE_BLOB) {
                    ClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) j, n, (short) (j + 1), n + 1);
                    anchor.setAnchorType(3);
                    patriarch.createPicture(anchor, workbook.addPicture(cursor.getBlob(j), HSSFWorkbook.PICTURE_TYPE_JPEG));
                } else {
                    String value = cursor.getString(j);
                    if (!TextUtils.isEmpty(value) && value.length() >= 32767) {
                        value = value.substring(0, 32766);
                    }
                    cellA.setCellValue(new HSSFRichTextString(value));
                }
            }
            n++;
            cursor.moveToNext();
        }
        cursor.close();
    }

    /**
     * get database all tables
     *
     * @return tables
     */
    private ArrayList<String> getAllTables(SQLiteDatabase database) {
        ArrayList<String> tables = new ArrayList<>();
        Cursor cursor = database.rawQuery("select name from sqlite_master where type='table' order by name", null);
        while (cursor.moveToNext()) {
            tables.add(cursor.getString(0));
        }
        cursor.close();
        return tables;
    }

    /**
     * get all columns from a table
     *
     * @param table database table
     * @return columns
     */
    private ArrayList<String> getTableColumns(SQLiteDatabase database, String table) {
        ArrayList<String> columns = new ArrayList<>();
        Cursor cursor = database.rawQuery("PRAGMA table_info(" + table + ")", null);
        while (cursor.moveToNext()) {
            columns.add(cursor.getString(1));
        }
        cursor.close();
        return columns;
    }

    /**
     * Callbacks for export events.
     */
    public interface ExportListener {
        void onStart();

        void onCompleted(String filePath);

        void onError(Exception e);
    }

}
