package com.liyu.sqlitetoexcel;

import android.content.ContentValues;
import android.content.Context;
import android.database.sqlite.SQLiteDatabase;
import android.os.Handler;
import android.os.Looper;
import android.text.TextUtils;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

/**
 * Excel to SQLite
 * Created by liyu on 2017/3/31.
 */

public class ExcelToSQLite {

    private static Handler handler = new Handler(Looper.getMainLooper());

    private Context mContext;
    private String dataBaseName;
    private SQLiteDatabase database;
    private String filePath;
    private String assetFileName;
    private String decryptKey;
    private String dateFormat;

    private SimpleDateFormat sdf;

    public static class Builder {
        private Context context;
        private String dataBaseName;
        private String filePath;
        private String assetFileName;
        private String decryptKey;
        private String dateFormat;

        public Builder(Context context) {
            this.context = context.getApplicationContext();
        }

        public ExcelToSQLite build() {
            if (TextUtils.isEmpty(dataBaseName)) {
                throw new IllegalArgumentException("Database name must not be null.");
            }
            return new ExcelToSQLite(context, dataBaseName, filePath, assetFileName, decryptKey, dateFormat);
        }

        public Builder setDataBase(String dataBaseName) {
            this.dataBaseName = dataBaseName;
            return this;
        }

        public Builder setDateFormat(String dateFormat) {
            this.dateFormat = dateFormat;
            return this;
        }

        public Builder setFilePath(String path) {
            this.filePath = path;
            this.assetFileName = null;
            if (TextUtils.isEmpty(this.dataBaseName)) {
                this.dataBaseName = context.getDatabasePath(new File(path).getName() + ".db").getAbsolutePath();
            }
            return this;
        }

        public Builder setDecryptKey(String decryptKey) {
            this.decryptKey = decryptKey;
            return this;
        }

        public Builder setAssetFileName(String name) {
            this.assetFileName = name;
            this.filePath = null;
            if (TextUtils.isEmpty(this.dataBaseName)) {
                this.dataBaseName = context.getDatabasePath(new File(name).getName() + ".db").getPath();
            }
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

    private ExcelToSQLite(Context context, String dataBaseName, String filePath, String assetFileName, String decryptKey, String dateFormat) {
        this.mContext = context;
        this.filePath = filePath;
        this.assetFileName = assetFileName;
        this.decryptKey = decryptKey;
        this.dataBaseName = dataBaseName;
        this.dateFormat = dateFormat;
        if (!TextUtils.isEmpty(dateFormat)) {
            sdf = new SimpleDateFormat(dateFormat);
        }

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
                                listener.onCompleted(dataBaseName);
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
            if (!TextUtils.isEmpty(decryptKey)) {
                Biff8EncryptionKey.setCurrentUserPassword("1234567");
            }
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
                    if (HSSFDateUtil.isCellDateFormatted(row.getCell(n))) {
                        if (sdf == null) {
                            values.put(columns.get(n), DateFormat.getDateTimeInstance().format(row.getCell(n).getDateCellValue()));
                        } else {
                            values.put(columns.get(n), sdf.format(row.getCell(n).getDateCellValue()));
                        }
                    } else {
                        String value = getRealStringValueOfDouble(row.getCell(n).getNumericCellValue());
                        values.put(columns.get(n), value);
                    }
                } else if (row.getCell(n).getCellType() == Cell.CELL_TYPE_STRING) {
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

    private static String getRealStringValueOfDouble(Double d) {
        String doubleStr = d.toString();
        boolean b = doubleStr.contains("E");
        int indexOfPoint = doubleStr.indexOf('.');
        if (b) {
            int indexOfE = doubleStr.indexOf('E');
            BigInteger xs = new BigInteger(doubleStr.substring(indexOfPoint
                    + BigInteger.ONE.intValue(), indexOfE));
            int pow = Integer.valueOf(doubleStr.substring(indexOfE
                    + BigInteger.ONE.intValue()));
            int xsLen = xs.toByteArray().length;
            int scale = xsLen - pow > 0 ? xsLen - pow : 0;
            doubleStr = String.format("%." + scale + "f", d);
        } else {
            java.util.regex.Pattern p = Pattern.compile(".0$");
            java.util.regex.Matcher m = p.matcher(doubleStr);
            if (m.find()) {
                doubleStr = doubleStr.replace(".0", "");
            }
        }
        return doubleStr;
    }

    /**
     * Callbacks for import events.
     */
    public interface ImportListener {
        void onStart();

        void onCompleted(String dataBaseName);

        void onError(Exception e);
    }

}
