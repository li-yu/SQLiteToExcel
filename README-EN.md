# SQLiteToExcel

[ ![Download](https://api.bintray.com/packages/li-yu/maven/SQLiteToExcel/images/download.svg) ](https://bintray.com/li-yu/maven/SQLiteToExcel/_latestVersion)

[中文版 README](README.md)

The SQLiteToExcel library integrates [Apache POI](http://poi.apache.org/) and some basic database query operations, making it easier to convert between SQLite and Excel.

From v1.0.5, **not support xlsx** format any more, because poi ooxml lib and other dependencies are so big( > 10 MB), and there are some strange problems on Android.

From v1.0.8, support custom SQL query.

## Version history
[Release Notes](https://github.com/li-yu/SQLiteToExcel/releases)

## How to use
#### 1. Add Gradle dependencies
``` Gradle
compile 'com.liyu.tools:sqlitetoexcel:1.0.10'
```

#### 2. SQLite -> Excel  [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/java/com/liyu/demo/MainActivity.java)
```java
new SQLiteToExcel
                .Builder(this)
                .setDataBase(databasePath) //Required. Tips: internal database path can be got by context.getDatabasePath("internal.db").getPath()
                .setTables(table1, table2) //Optional, if null, all tables will be export. 
                .setOutputPath(outoutPath) //Optional, if null, default output path is app ExternalFilesDir. 
                .setOutputFileName("test.xls") //Optional, if null, default output file name is xxx.db.xls
                .setEncryptKey("1234567") //Optional, if you want to encrypt the output file.
                .setProtectKey("9876543") //Optional, if you want to set the sheet read only.
                .start(ExportListener); // or .start() for synchronous method.
```

Custom SQL：
```java
new SQLiteToExcel
                .Builder(this)
                .setDataBase(databasePath)
                .setSQL("select name as USER_NAME, price as USER_PRICE from user where name like '%FRANK%'")
                .start(ExportListener);
```

#### 3. Excel -> SQLite [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/java/com/liyu/demo/MainActivity.java)
```java
new ExcelToSQLite
                .Builder(this)
                .setDataBase(databasePath) // Optional, default is "*.xls.db" in internal database path.
                .setAssetFileName("user.xls") // if it is a asset file.
                .setFilePath("/storage/doc/user.xls") // if it is a normal file.
                .setDecryptKey("1234567") // Optional, if need to decrypt the file.
                .setDateFormat("yyyy-MM-dd HH:mm:ss") // Optional, if need to format date cell.
                .start(ImportListener); // or .start() for synchronous method.
```

#### 4. Thanks (how to support xlsx?)
- [https://github.com/centic9/poi-on-android](https://github.com/centic9/poi-on-android)
- [https://github.com/FasterXML/aalto-xml](https://github.com/FasterXML/aalto-xml)
- [https://github.com/johnrengelman/shadow](https://github.com/johnrengelman/shadow)

#### 5. Precautions
* When read or write external file, Android 6.0 and above need to deal with run-time permissions.
* When convert Excel to SQLite, take excel sheet first row as the database table column name, please refer to the style [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/assets/user.xls).
* Currently only blob field is supported as a picture, because I do not know whether byte [] is a file or a picture.
* ~~The database files must be located under ```/data/data/package name/databases/` ``, and are usually located in this directory.~~

## About me
* Email: [me@liyuyu.cn](mailto:me@liyuyu.cn)
* Weibo: [@呵呵小小鱼](http://weibo.com/u/1241167880)
