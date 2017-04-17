# SQLiteToExcel

[ ![Download](https://api.bintray.com/packages/li-yu/maven/SQLiteToExcel/images/download.svg) ](https://bintray.com/li-yu/maven/SQLiteToExcel/_latestVersion)

[English README](README-EN.md)

SQLiteToExcel 库整合了 [Apache POI](http://poi.apache.org/) 和一些基本的数据库查询操作，使得 SQLite 和 Excel 之间相互转换更加便捷。

## 更新历史
2017-04-17 ： v1.0.4 
- 支持 xlsx 格式输入和输出，感谢 @DearZack 反馈 [issue](https://github.com/li-yu/SQLiteToExcel/issues/2)
- 为了支持 xlsx，方法数已经爆表，请开启 ``multiDexEnabled``，没有特殊需求，建议还是使用仅支持 xls 的 v1.0.3 版本
- 修复了 excel 导入时空白行可能引起报错的 bug

2017-03-31 ： v1.0.3 
- 新增 Excel 导入 SQLite 数据库的功能

2017-03-28 ： v1.0.2 
- 上传到 JCenter

2017-03-24 ： v1.0.2 
- 解决 blob 字段导出报错的 bug
- 目前仅支持 blob 字段导出为图片
- Apache POI 版本同步更新到 v3.15

2015-12-25 ： v1.0.1 
- 可以设置导出目录，默认为内部SD卡根目录
- Apache POI 版本同步更新到 v3.13

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

#### 5.注意事项
* Excel 导入 SQLite 时，默认取 excel 中 sheet 的**第一行**作为数据库表的列名，样式请参考 [demo](https://github.com/li-yu/SQLiteToExcel/blob/master/app/src/main/assets/user.xls)。
* 目前仅支持 blob 字段导出为图片，因为我也不知道 byte[] 是文件还是图片。
* 数据库文件须位于```/data/data/包名/databases/```下，一般都是位于这个目录下。

## 关于我
* Email: [me@liyuyu.cn](mailto:me@liyuyu.cn)
* Weibo: [@呵呵小小鱼](http://weibo.com/u/1241167880)
