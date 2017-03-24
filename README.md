# SQLiteToExcel
SQLiteToExcel 库整合了 [Apache POI](http://poi.apache.org/) 和一些基本的数据库查询操作，使得生成 excel 文件更加便捷。

## 更新历史
2017-03-24 ： v1.0.2 

- 解决 blob 字段导出报错的 bug
- 目前仅支持 blob 字段导出为图片
- Apache POI 版本同步更新到 v3.15


2015-12-25 ： v1.0.1 

- 可以设置导出目录，默认为内部SD卡根目录
- Apache POI 版本同步更新到 v3.13

## 主要功能
* 1.导出单个表
* 2.导出多个表
* 3.导出所有表

## 如何使用
#### 1.添加 SD 卡读写权限到 AndroidManifest.xml
```xml
<uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" />
```
#### 2.下载 Jar 文件作为 libs 添加到工程中
[SqliteToExcel-v1.0.2.jar](https://github.com/li-yu/SQLiteToExcel/raw/master/SqliteToExcel-v1.0.2.jar)
#### 3.示例代码
* 初始化（默认导出路径为外部 SD 卡根目录）
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
#### 4.注意事项
* 目前仅支持 blob 字段导出为图片，因为我也不知道 byte[] 是文件还是图片。
* 数据库文件须位于```/data/data/包名/databases/```下。一般都是位于这个目录下。
* ~~excel文件生成路径为：```Environment.getExternalStorageDirectory()```，即外部SD卡根目录。下一个版本会修改代码，可以指定生成的路径。~~

## 关于我
* Email: [me@liyuyu.cn](mailto:me@liyuyu.cn)
* Weibo: [@呵呵小小鱼](http://weibo.com/u/1241167880)
