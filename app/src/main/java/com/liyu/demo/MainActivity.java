package com.liyu.demo;

import android.content.res.Resources;
import android.database.Cursor;
import android.database.sqlite.SQLiteDatabase;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.support.annotation.DrawableRes;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.TextView;

import com.liyu.sqlitetoexcel.ExcelToSqlite;
import com.liyu.sqlitetoexcel.SqliteToExcel;

import java.io.ByteArrayOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

public class MainActivity extends AppCompatActivity {

    private TextView tv;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        tv = (TextView) findViewById(R.id.tv);

        User user = new User();
        user.setName("呵呵小鱼");
        user.setPrice(19.89f);
        user.setCover(getResByte(R.drawable.yuyu));
        user.save();

    }

    public void onExport(View v) {
        SqliteToExcel ste = new SqliteToExcel(this, "ste_db.db");
        ste.startExportAllTables("test.xls", new SqliteToExcel.ExportListener() {
            @Override
            public void onStart() {
                tv.append(makeLog("\n" +
                        "Export start--->"));
            }

            @Override
            public void onCompleted(String filePath) {
                tv.append(makeLog("\n" +
                        "Export completed--->" + filePath));
            }

            @Override
            public void onError(Exception e) {
                tv.append(makeLog("\n" +
                        "Export error--->" + e.toString()));

            }
        });
    }

    public void onImport(View v) {
        ExcelToSqlite ets = new ExcelToSqlite(this, "user.db");
        ets.startFromAsset("user.xls", new ExcelToSqlite.ImportListener() {
            @Override
            public void onStart() {
                tv.append(makeLog("\n" +
                        "Import start--->"));
            }

            @Override
            public void onCompleted(String dbName) {
                tv.append(makeLog("\n" +
                        "Import completed--->"));
                showDbMsg(dbName);
            }

            @Override
            public void onError(Exception e) {
                tv.append(makeLog("\n" +
                        "Import error--->" + e.toString()));
            }
        });
    }

    private void showDbMsg(String dbName) {
        SQLiteDatabase database;
        try {
            database = SQLiteDatabase.openOrCreateDatabase(MainActivity.this.getDatabasePath(dbName).getAbsolutePath(), null);
            Cursor cursor = database.rawQuery("select name from sqlite_master where type='table' order by name", null);
            while (cursor.moveToNext()) {
                tv.append("\nNew tables is : " + cursor.getString(0) + "  ");
                Cursor cursor2 = database.rawQuery("select * from " + cursor.getString(0), null);
                while (cursor2.moveToNext()) {
                    tv.append("\n");
                    for (int i = 0; i < cursor2.getColumnCount(); i++) {
                        tv.append(cursor2.getString(i) + "  ");
                    }
                }
                cursor2.close();
            }
            cursor.close();
            database.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String makeLog(String log) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date now = new Date();
        return log + " " + sdf.format(now);
    }

    private byte[] getResByte(@DrawableRes int id) {
        Resources res = getResources();
        Bitmap bmp = BitmapFactory.decodeResource(res, id);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        bmp.compress(Bitmap.CompressFormat.PNG, 100, baos);
        return baos.toByteArray();
    }
}
