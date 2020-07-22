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
import android.support.annotation.NonNull;
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
    private String sql;
    private String sheetName;

    private SQLiteDatabase database;
    private Workbook workbook;

    public static class Builder {
        private String dataBaseName;
        private String filePath;
        private String fileName;
        private String protectKey;
        private String encryptKey;
        private List<String> tables;
        private String sql;
        private String sheetName;
        private Context context;

        public Builder(Context context) {
            this.context = context;
        }

        public SQLiteToExcel build() {
            if (TextUtils.isEmpty(dataBaseName)) {
                throw new IllegalArgumentException("Database name must not be null.");
            }
            if (TextUtils.isEmpty(filePath)) {
                this.filePath = context.getExternalFilesDir(null).getPath();
            }
            if (TextUtils.isEmpty(fileName)) {
                this.fileName = new File(dataBaseName).getName() + ".xls";
            }
            return new SQLiteToExcel(tables, protectKey, encryptKey, fileName, dataBaseName, filePath, sql, sheetName);
        }

        public Builder setDataBase(String dataBaseName) {
            this.dataBaseName = dataBaseName;
            return this;
        }

        /**
         * @param fileName
         * @return Builder
         * @deprecated Use {@link #setOutputFileName(String fileName)} instead.
         */
        @Deprecated
        public Builder setFileName(String fileName) {
            return setOutputFileName(fileName);
        }

        public Builder setOutputFileName(String fileName) {
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

        /**
         * @param path
         * @return Builder
         * @deprecated Use {@link #setOutputPath(String path)} instead.
         */
        @Deprecated
        public Builder setPath(String path) {
            return setOutputPath(path);
        }

        public Builder setOutputPath(String path) {
            this.filePath = path;
            return this;
        }

        public Builder setSQL(@NonNull String sheetName, @NonNull String sql) {
            this.sql = sql;
            this.sheetName = sheetName;
            return this;
        }

        public Builder setSQL(@NonNull String sql) {
            return setSQL("Sheet1", sql);
        }

        public String start() throws Exception {
            final SQLiteToExcel sqliteToExcel = build();
            return sqliteToExcel.start();
        }

        public void start(ExportListener listener) {
            final SQLiteToExcel sqliteToExcel = build();
            sqliteToExcel.start(listener);
        }
    }

    /**
     * import Tables task
     *
     * @return output file path
     */
    public String start() throws Exception {
        try {
            if (tables == null || tables.size() == 0) {
                tables = getTablesName(database);
            }
            return exportTables(tables, fileName);
        } catch (Exception e) {
            if (database != null && database.isOpen()) {
                database.close();
            }
            throw e;
        }
    }

    /**
     * importTables task with a listener
     *
     * @param listener callback
     */
    public void start(final ExportListener listener) {
        if (listener != null) {
            listener.onStart();
        }
        new Thread(new Runnable() {

            @Override
            public void run() {
                try {
                    if (tables == null || tables.size() == 0) {
                        tables = getTablesName(database);
                    }
                    final String finalFilePath = exportTables(tables, fileName);
                    if (listener != null) {
                        handler.post(new Runnable() {
                            @Override
                            public void run() {
                                listener.onCompleted(finalFilePath);
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

    private SQLiteToExcel(List<String> tables, String protectKey, String encryptKey, String fileName,
                          String dataBaseName, String filePath, String sql, String sheetName) {
        this.protectKey = protectKey;
        this.encryptKey = encryptKey;
        this.fileName = fileName;
        this.filePath = filePath;
        this.sql = sql;
        this.sheetName = sheetName;
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
            throw new IllegalArgumentException("File name is null or unsupported file format!");
        }
        if (TextUtils.isEmpty(sql)) {
            for (int i = 0; i < tables.size(); i++) {
                Sheet sheet = workbook.createSheet(tables.get(i));
                String sqlAll = "select * from " + tables.get(i);
                fillSheet(sqlAll, sheet);
                if (!TextUtils.isEmpty(protectKey)) {
                    sheet.protectSheet(protectKey);
                }
            }
        } else {
            Sheet sheet = workbook.createSheet(sheetName);
            fillSheet(sql, sheet);
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
     * @param sql   query sql
     * @param sheet target sheet
     */
    private void fillSheet(String sql, Sheet sheet) {
        Drawing patriarch = sheet.createDrawingPatriarch();
        Cursor cursor = database.rawQuery(sql, null);
        cursor.moveToFirst();
        final int columnsCount = cursor.getColumnCount();
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columnsCount; i++) {
            Cell cellA = headerRow.createCell(i);
            cellA.setCellValue(new HSSFRichTextString("" + cursor.getColumnNames()[i]));
        }
        int n = 1;
        while (!cursor.isAfterLast()) {
            Row rowA = sheet.createRow(n);
            for (int j = 0; j < columnsCount; j++) {
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
    private List<String> getTablesName(SQLiteDatabase database) {
        List<String> tables = new ArrayList<>();
        Cursor cursor = database.rawQuery("select name from sqlite_master where type='table' order by name", null);
        while (cursor.moveToNext()) {
            tables.add(cursor.getString(0));
        }
        cursor.close();
        return tables;
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
