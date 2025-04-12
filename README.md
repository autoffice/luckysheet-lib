luckysheet-lib
======================
luckysheet-lib是Luckysheet的Java库，包括excel导入luckysheet和luckysheet导出为xlsx格式的excel文件。

## 使用方法
```java
import com.helloaldis.autoffice.luckyshee.LuckysheetConverter;
...
        
LuckysheetConverter.luckysheetToExcel("/path/luckysheet.json", "/path/excel.xlsx");
LuckysheetConverter.excelToLuckySheetFile("/path/excel.xlsx", "/path/luckysheet.json");
```