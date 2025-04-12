luckysheet-lib
======================
luckysheet-lib是Luckysheet的Java库，包括excel导入luckysheet和luckysheet导出为xlsx格式的excel文件。

## 使用方法
```java
import com.helloaldis.autoffice.luckyshee.LuckysheetConverter;

public class Test {
    public static void main(String[] args) {
        // 将luckysheet json文件转为excel
        LuckysheetConverter.luckysheetToExcel("/path/luckysheet.json", "/path/excel.xlsx");
        
        // 将excel转为luckysheet json文件 
        LuckysheetConverter.excelToLuckySheetFile("/path/excel.xlsx", "/path/luckysheet.json");
    }
}
```