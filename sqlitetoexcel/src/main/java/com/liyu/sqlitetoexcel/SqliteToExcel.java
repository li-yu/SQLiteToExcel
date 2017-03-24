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
 * 导出SQLite数据库，生成excel文件
 * <p>
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

    /**
     * 构造函数
     *
     * @param context 上下文
     * @param dbName  数据库名称
     */
    public SqliteToExcel(Context context, String dbName) {
        this(context, dbName, Environment.getExternalStorageDirectory().toString() + File.separator);
    }

    /**
     * @param context    上下文
     * @param dbName     数据库名称
     * @param exportPath 导出文件的路径（不包括文件名）
     */
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

    /**
     * 获取数据库中所有表名
     *
     * @return
     */
    private ArrayList<String> getAllTables() {
        ArrayList<String> tables = new ArrayList<>();
        Cursor cursor = database.rawQuery("select name from sqlite_master where type='table' order by name", null);
        while (cursor.moveToNext()) {
            tables.add(cursor.getString(0));
        }
        cursor.close();
        return tables;
    }

    /**
     * 获取一个表的所有列名
     *
     * @param table
     * @return
     */
    private ArrayList<String> getColumns(String table) {
        ArrayList<String> columns = new ArrayList<>();
        Cursor cursor = database.rawQuery("PRAGMA table_info(" + table + ")", null);
        while (cursor.moveToNext()) {
            columns.add(cursor.getString(1));
        }
        cursor.close();
        return columns;
    }

    /**
     * 导出数据库中多个表到excel文件
     *
     * @param fileName 生成的excel文件名
     * @param tables   表名
     * @throws Exception
     */
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

    /**
     * 开始导出单个表的任务
     *
     * @param table    表名
     * @param fileName 生成的excel文件名
     * @param listener 任务监听器
     */
    public void startExportSingleTable(final String table, final String fileName, ExportListener listener) {
        List<String> tables = new ArrayList<>();
        tables.add(table);
        startExportTables(tables, fileName, listener);
    }

    /**
     * 开始导出所有表的的任务
     *
     * @param fileName 生成的excel文件名
     * @param listener 任务监听器
     */
    public void startExportAllTables(final String fileName, ExportListener listener) {
        ArrayList<String> tables = getAllTables();
        startExportTables(tables, fileName, listener);
    }

    /**
     * 开始导出多个表的任务
     *
     * @param tables   表名
     * @param fileName 生成的excel文件名
     * @param listener 任务监听器
     */
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

    /**
     * 创建excel的sheet
     *
     * @param table 数据库表名
     * @param sheet 工作薄中的sheet
     */
    private void createSheet(String table, HSSFSheet sheet) {
        HSSFRow rowA = sheet.createRow(0);
        ArrayList<String> columns = getColumns(table);
        for (int i = 0; i < columns.size(); i++) {
            HSSFCell cellA = rowA.createCell(i);
            cellA.setCellValue(new HSSFRichTextString("" + columns.get(i)));
        }
        insertItemToSheet(table, sheet, columns);
    }

    /**
     * 插入数据到工作薄中的sheet
     *
     * @param table   数据库表名
     * @param sheet   工作薄中的sheet
     * @param columns 所有列
     */
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
                    //插入图片
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

    /**
     * 任务监听器接口
     *
     * @author yu.li
     */
    public interface ExportListener {
        void onStart();

        void onCompleted(String filePath);

        void onError(Exception e);
    }

}
