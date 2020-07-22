package com.liyu.demokt

import android.app.ProgressDialog
import android.os.Bundle
import android.os.Environment
import android.support.v7.app.AppCompatActivity
import android.view.View
import android.widget.Toast
import com.liyu.sqlitetoexcel.kt.SQLiteToExcel
import java.io.File

class MainActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
    }

    fun onClick(view: View) {
        val progressDialog = ProgressDialog(this)
        progressDialog.setMessage("导出中...")
        val dbpath = (Environment.getExternalStorageDirectory().path
                + File.separator + "citys.db")

        SQLiteToExcel(this).start {
            databasePath = dbpath
            tables = arrayOf("citys")
            outputPath = Environment.getExternalStorageDirectory().path
            outputFileName = "2020citys.xls"
            encryptKey = "12345678"
            protectKey = "87654321"
            onStart {
                progressDialog.show()
            }
            onCompleted {
                progressDialog.dismiss()
                Toast.makeText(this@MainActivity, "导出成功!$it", Toast.LENGTH_SHORT).show()
            }
            onError {
                Toast.makeText(this@MainActivity, "$it", Toast.LENGTH_SHORT).show()
            }
        }
    }
}
