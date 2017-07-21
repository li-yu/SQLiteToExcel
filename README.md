# SQLiteToExcel

[ ![Download](https://api.bintray.com/packages/li-yu/maven/SQLiteToExcel/images/download.svg) ](https://bintray.com/li-yu/maven/SQLiteToExcel/_latestVersion)

[English README](README-EN.md)

SQLiteToExcel 库整合了 [Apache POI](http://poi.apache.org/) 和一些基本的数据库操作，使得 SQLite 和 Excel 之间相互转换更加便捷。

从 v1.0.5 版本开始，不再支持 **xlsx 格式**，因为 poi ooxml 库和其他一些相关的依赖太大了，体积超过了 10MB，同时开发过程中也发现 poi 对于 Android 支持不够全面，放弃 xlsx 也是挺无奈的，个人觉得 xls 格式对于我们来说已经够用。

## 更新历史
[Release Notes](https://github.com/li-yu/SQLiteToExcel/releases)

## 如何使用
#### 1. 添加 Gradle 依赖
``` Gradle
compile 'com.liyu.tools:sqlitetoexcel:1.0.5'
```

#### 2. SQLite -> Excel 示例代码（具体示例可参考 [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/java/com/liyu/demo/MainActivity.java) 工程）
```java
new SQLiteToExcel
                .Builder(this)
                .setDataBase(databasePath) //必须。 小提示：内部数据库可以通过 context.getDatabasePath("internal.db").getPath() 获取。
                .setTables(table1, table2) //可选, 如果不设置，则默认导出全部表。
                .setPath(outoutPath) //可选, 如果不设置，默认输出路径为 app ExternalFilesDir。
                .setFileName("test.xls") //可选, 如果不设置，输出的文件名为 xxx.db.xls。
                .setEncryptKey("1234567") //可选，可对导出的文件进行加密。
                .setProtectKey("9876543") //可选，可对导出的表格进行只读的保护。
                .start(ExportListener); // 或者使用 .start() 同步方法。
```

#### 3. Excel -> SQLite 示例代码（具体示例可参考 [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/java/com/liyu/demo/MainActivity.java) 工程）
```java
new ExcelToSQLite
                .Builder(this)
                .setDataBase(databasePath) //必须。
                .setAssetFileName("user.xls") // 如果文件在 asset 目录。
                .setFilePath("/storage/doc/user.xls") // 如果文件在其他目录。
                .start(ImportListener); // 或者使用 .start() 同步方法。
```

#### 4. 感谢（如何支持 xlsx 可以参考以下仓库）
- [https://github.com/centic9/poi-on-android](https://github.com/centic9/poi-on-android)
- [https://github.com/FasterXML/aalto-xml](https://github.com/FasterXML/aalto-xml)
- [https://github.com/johnrengelman/shadow](https://github.com/johnrengelman/shadow)

#### 5. 注意事项
* 读写外部文件时，Android 6.0 及以上版本需处理运行时权限。
* Excel 导入 SQLite 时，默认取 excel 中 sheet 的**第一行**作为数据库表的列名，样式请参考 [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/assets/user.xls)。
* 目前仅支持 blob 字段导出为图片，因为我也不知道 byte[] 是文件还是图片。
* ~~数据库文件须位于```/data/data/包名/databases/```下，一般都是位于这个目录下。~~

## 联系我
* Email: [me@liyuyu.cn](mailto:me@liyuyu.cn)
* Weibo: [@呵呵小小鱼](http://weibo.com/u/1241167880)
