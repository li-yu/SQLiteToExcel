package com.liyu.sqlitetoexcel.kt

import android.content.Context
import java.lang.Exception

class SQLiteToExcel(val context: Context) {

    var databasePath: String? = null
    var tables: Array<String> = arrayOf("")
    var outputPath: String? = null
    var outputFileName: String? = null
    var encryptKey: String? = null
    var protectKey: String? = null

    private var onCompleted: ((String) -> Unit)? = null
    private var onError: ((Throwable) -> Unit)? = null
    private var onStart: (() -> Unit)? = null

    fun start(block: SQLiteToExcel.() -> Unit) {
        block()
        com.liyu.sqlitetoexcel.SQLiteToExcel.Builder(context)
                .setDataBase(databasePath)
                .setTables(*tables)
                .setOutputPath(outputPath)
                .setOutputFileName(outputFileName)
                .setEncryptKey(encryptKey)
                .setProtectKey(protectKey)
                .start(object : com.liyu.sqlitetoexcel.SQLiteToExcel.ExportListener {
                    override fun onStart() {
                        onStartInner()
                    }

                    override fun onCompleted(filePath: String) {
                        onCompletedInner(filePath)
                    }

                    override fun onError(e: Exception) {
                        onErrorInner(e)
                    }
                })
    }


    fun onCompleted(block: (String) -> Unit) {
        onCompleted = block
    }

    fun onError(block: (Throwable) -> Unit) {
        onError = block
    }

    fun onStart(block: () -> Unit) {
        onStart = block
    }

    private fun onCompletedInner(path: String) {
        onCompleted?.invoke(path)
    }

    private fun onErrorInner(exception: Throwable) {
        onError?.invoke(exception)
    }

    private fun onStartInner() {
        onStart?.invoke()
    }
}