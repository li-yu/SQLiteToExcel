# SQLiteToExcel

[ ![Download](https://api.bintray.com/packages/li-yu/maven/SQLiteToExcel/images/download.svg) ](https://bintray.com/li-yu/maven/SQLiteToExcel/_latestVersion)

[English README](README-EN.md)

SQLiteToExcel 库整合了 [Apache POI](http://poi.apache.org/) 和一些基本的数据库查询操作，使得 SQLite 和 Excel 之间相互转换更加便捷。

## 更新历史
[Release Notes](https://github.com/li-yu/SQLiteToExcel/releases)

## 如何使用
#### 1.添加 Gradle 依赖或者下载 Jar 文件作为 libs 添加到工程中
``` Gradle
compile 'com.liyu.tools:sqlitetoexcel:1.0.4'
```
[SqliteToExcel-v1.0.4.jar](https://github.com/li-yu/SQLiteToExcel/releases)
#### 2.添加 SD 卡读写权限到 AndroidManifest.xml（Android 6.0 及以上需要处理运行时权限）
```xml
<uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" />
```

#### 3.SQLite -> Excel 示例代码（具体示例可参考 [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/java/com/liyu/demo/MainActivity.java) 工程）
* 初始化（默认导出路径为外部 SD 卡根目录 ```Environment.getExternalStorageDirectory()```）
```java
SqliteToExcel ste = new SqliteToExcel(this, "helloworld.db");
```
或（指定导出的根目录）
```java
SqliteToExcel ste = new SqliteToExcel(this, "helloworld.db", "/mnt/sdcard/myfiles/");
```
* 导出单个表到 excel
```java
ste.startExportSingleTable(String table, String fileName, ExportListener listener);
```
* 导出多个表到 excel
```java
ste.startExportTables(List<String> tables, String fileName, ExportListener listener);
```
* 导出所有表到 excel
```java
ste.startExportAllTables(String fileName, ExportListener listener);
```
* 任务监听器接口
```java
public interface ExportListener {
        void onStart();

        void onCompleted(String filePath);

        void onError(Exception e);
    }
```

#### 4.Excel -> SQLite 示例代码（具体示例可参考 [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/java/com/liyu/demo/MainActivity.java) 工程）
* 初始化
```java
ExcelToSqlite ets = new ExcelToSqlite(this, "user.db");
```
* 从 assets 目录传入 excel 文件
```java
ets.startFromAsset(String assetFileName, ImportListener listener);
```
* 以 File 形式传入任意 excel 文件
```java
ets.startFromFile(File file, ImportListener listener);
```
* 任务监听器接口
```java
public interface ImportListener {
        void onStart();

        void onCompleted(String dbName);

        void onError(Exception e);
    }
```

#### 5.感谢
- [https://github.com/centic9/poi-on-android](https://github.com/centic9/poi-on-android)
- [https://github.com/FasterXML/aalto-xml](https://github.com/FasterXML/aalto-xml)
- [https://github.com/johnrengelman/shadow](https://github.com/johnrengelman/shadow)

#### 6.注意事项
* Excel 导入 SQLite 时，默认取 excel 中 sheet 的**第一行**作为数据库表的列名，样式请参考 [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/assets/user.xls)。
* 目前仅支持 blob 字段导出为图片，因为我也不知道 byte[] 是文件还是图片。
* 数据库文件须位于```/data/data/包名/databases/```下，一般都是位于这个目录下。

## 关于我
* Email: [me@liyuyu.cn](mailto:me@liyuyu.cn)
* Weibo: [@呵呵小小鱼](http://weibo.com/u/1241167880)
