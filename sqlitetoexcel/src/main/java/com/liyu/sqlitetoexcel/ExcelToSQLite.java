package com.liyu.sqlitetoexcel;

import android.content.ContentValues;
import android.content.Context;
import android.database.sqlite.SQLiteDatabase;
import android.os.Handler;
import android.os.Looper;
import android.text.TextUtils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Excel to SQLite
 * Created by liyu on 2017/3/31.
 */

public class ExcelToSQLite {

    private static Handler handler = new Handler(Looper.getMainLooper());

    private Context mContext;
    private SQLiteDatabase database;
    private String filePath;
    private String assetFileName;

    public static class Builder {
        private Context context;
        private String dataBaseName;
        private String filePath;
        private String assetFileName;

        public Builder(Context context) {
            this.context = context.getApplicationContext();
        }

        public ExcelToSQLite build() {
            if (TextUtils.isEmpty(dataBaseName)) {
                throw new IllegalArgumentException("Database name must not be null.");
            }
            return new ExcelToSQLite(context, dataBaseName, filePath, assetFileName);
        }

        public Builder setDataBase(String dataBaseName) {
            this.dataBaseName = dataBaseName;
            return this;
        }

        public Builder setFilePath(String path) {
            this.filePath = path;
            this.assetFileName = null;
            return this;
        }

        public Builder setAssetFileName(String name) {
            this.assetFileName = name;
            this.filePath = null;
            return this;
        }

        public void start() {
            final ExcelToSQLite excelToSqlite = build();
            excelToSqlite.start();
        }

        public void start(ImportListener listener) {
            final ExcelToSQLite excelToSqlite = build();
            excelToSqlite.start(listener);
        }

    }

    private ExcelToSQLite(Context mContext, String dataBaseName, String filePath, String assetFileName) {
        this.mContext = mContext;
        this.filePath = filePath;
        this.assetFileName = assetFileName;

        try {
            database = SQLiteDatabase.openOrCreateDatabase(dataBaseName, null);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * start task
     */
    public boolean start() {
        if (TextUtils.isEmpty(filePath) && TextUtils.isEmpty(assetFileName)) {
            throw new IllegalArgumentException("Asset file or external file name must not be null.");
        }
        try {
            if (TextUtils.isEmpty(filePath)) {
                return importTables(mContext.getAssets().open(assetFileName), assetFileName);
            } else {
                return importTables(new FileInputStream(filePath), filePath);
            }
        } catch (Exception e) {
            if (database != null && database.isOpen()) {
                database.close();
            }
            return false;
        }
    }

    /**
     * start task with a listener
     *
     * @param listener
     */
    public void start(final ImportListener listener) {
        if (TextUtils.isEmpty(filePath) && TextUtils.isEmpty(assetFileName)) {
            throw new IllegalArgumentException("Asset file or external file name must not be null.");
        }
        if (listener != null) {
            listener.onStart();
        }
        new Thread(new Runnable() {
            @Override
            public void run() {
                try {
                    if (TextUtils.isEmpty(filePath)) {
                        importTables(mContext.getAssets().open(assetFileName), assetFileName);
                    } else {
                        importTables(new FileInputStream(filePath), filePath);
                    }
                    if (listener != null) {
                        handler.post(new Runnable() {
                            @Override
                            public void run() {
                                listener.onCompleted(true);
                            }
                        });
                    }
                } catch (final Exception e) {
                    if (database != null && database.isOpen()) {
                        database.close();
                    }
                    if (listener != null) {
                        handler.post(new Runnable() {
                            @Override
                            public void run() {
                                listener.onError(e);
                            }
                        });
                    }
                }
            }
        }).start();
    }

    /**
     * core code
     *
     * @param stream   asset stream or file stream
     * @param fileName origin file name
     * @throws Exception
     */
    private boolean importTables(InputStream stream, String fileName) throws Exception {
        Workbook workbook;
        if (fileName.toLowerCase().endsWith(".xls")) {
            workbook = new HSSFWorkbook(stream);
        } else {
            throw new UnsupportedOperationException("Unsupported file format!");
        }
        stream.close();
        int sheetNumber = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetNumber; i++) {
            createTable(workbook.getSheetAt(i));
        }
        database.close();
        return true;
    }

    /**
     * create table by sheet
     *
     * @param sheet
     */
    private void createTable(Sheet sheet) {
        StringBuilder createTableSql = new StringBuilder("CREATE TABLE IF NOT EXISTS ");
        createTableSql.append(sheet.getSheetName());
        createTableSql.append("(");
        Iterator<Row> rit = sheet.rowIterator();
        Row rowHeader = rit.next();
        List<String> columns = new ArrayList<>();
        for (int i = 0; i < rowHeader.getPhysicalNumberOfCells(); i++) {
            createTableSql.append(rowHeader.getCell(i).getStringCellValue());
            if (i == rowHeader.getPhysicalNumberOfCells() - 1) {
                createTableSql.append(" TEXT");
            } else {
                createTableSql.append(" TEXT,");
            }
            columns.add(rowHeader.getCell(i).getStringCellValue());
        }
        createTableSql.append(")");
        database.execSQL(createTableSql.toString());
        while (rit.hasNext()) {
            Row row = rit.next();
            ContentValues values = new ContentValues();
            for (int n = 0; n < row.getPhysicalNumberOfCells(); n++) {
                if (row.getCell(n) == null) {
                    continue;
                }
                if (row.getCell(n).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    values.put(columns.get(n), row.getCell(n).getNumericCellValue());
                } else {
                    values.put(columns.get(n), row.getCell(n).getStringCellValue());
                }
            }
            if (values.size() == 0)
                continue;
            long result = database.insert(sheet.getSheetName(), null, values);
            if (result < 0) {
                throw new RuntimeException("Insert value failed!");
            }
        }
    }

    /**
     * Callbacks for import events.
     */
    public interface ImportListener {
        void onStart();

        void onCompleted(boolean result);

        void onError(Exception e);
    }

}
