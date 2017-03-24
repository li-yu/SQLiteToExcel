package com.liyu.demo;

import android.content.res.Resources;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.support.annotation.DrawableRes;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.text.format.DateUtils;
import android.view.View;
import android.widget.TextView;

import com.liyu.sqlitetoexcel.SqliteToExcel;

import java.io.ByteArrayOutputStream;

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

    public void onClick(View v) {
        SqliteToExcel ste = new SqliteToExcel(this, "ste_db.db");
        ste.startExportAllTables("test.xls", new SqliteToExcel.ExportListener() {
            @Override
            public void onStart() {
                tv.append(makeLog("\n" +
                        "start--->"));
            }

            @Override
            public void onCompleted(String filePath) {
                tv.append(makeLog("\n" +
                        "completed--->" + filePath));
            }

            @Override
            public void onError(Exception e) {
                tv.append(makeLog("\n" +
                        "error--->" + e.toString()));
            }
        });
    }

    private String makeLog(String log) {
        return log + DateUtils.formatDateTime(MainActivity.this, System.currentTimeMillis(), DateUtils.FORMAT_SHOW_YEAR |
                DateUtils.FORMAT_SHOW_DATE |
                DateUtils.FORMAT_SHOW_WEEKDAY |
                DateUtils.FORMAT_SHOW_TIME);
    }

    private byte[] getResByte(@DrawableRes int id) {
        Resources res = getResources();
        Bitmap bmp = BitmapFactory.decodeResource(res, id);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        bmp.compress(Bitmap.CompressFormat.PNG, 100, baos);
        return baos.toByteArray();
    }
}
