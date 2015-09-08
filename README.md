# SQLiteToExcel
SQLiteToExcel库整合了POI和一些基本的数据库查询操作，使得生成excel文件更加便捷。

##主要功能
* 1.导出单个表
* 2.导出所有表

##如何使用
####1.添加SD卡读写权限到AndroidManifest.xml
```xml
<uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" />
```
####2.下载Jar文件作为libs添加到工程中
[SqliteToExcel-v1.0.0.jar](https://github.com/li-yu/SQLiteToExcel/blob/master/SqliteToExcel-v1.0.0.jar?raw=true)
####3.示例代码
* 导出单个表到excel
```java
SqliteToExcel ste = new SqliteToExcel(this, "helloworld.db");
ste.startExportSingleTable("table1", "a.xls", new ExportListener() {
			
	@Override
	public void onStart() {
		
	}
			
	@Override
	public void onError() {
		
	}
			
	@Override
	public void onComplete() {
		
	}
});
```
* 导出所有表到excel
```java
ste.startExportAllTables("b.xls", new ExportListener() {
			
	@Override
	public void onStart() {
		
	}
			
	@Override
	public void onError() {
		
	}
			
	@Override
	public void onComplete() {
		
	}
});
```
####4.注意事项
* 数据库文件须位于```/data/data/com.xxx.xxx/databases/```下。下一个版本会修改代码，可以指定数据库路径
* excel文件生成路径为：```Environment.getExternalStorageDirectory()```，即外部SD卡根目录。下一个版本会修改代码，可以指定生成的路径

##关于我
* Email:[me@liyuyu.cn](mailto:me@liyuyu.cn)
* Weibo:[@呵呵小小鱼](http://weibo.com/u/1241167880)
