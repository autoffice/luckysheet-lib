luckysheet-lib
======================
[![Java CI](https://github.com/autoffice/luckysheet-lib/actions/workflows/ci.yml/badge.svg)](https://github.com/autoffice/luckysheet-lib/actions/workflows/ci.yml)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)

luckysheet-lib是Luckysheet的Java库，包括excel导入luckysheet和luckysheet导出为xlsx格式的excel文件。

## 使用方法
pom.xml引入luckysheet-lib依赖
```xml
            <dependency>
                <groupId>io.github.autoffice</groupId>
                <artifactId>luckysheet-lib</artifactId>
                <version>1.0.0</version>
            </dependency>
```

使用`LuckysheetConverter`类中对应的导入、导出方法即可，例如：
```java
import io.github.autoffice.luckysheet.LuckysheetConverter;

public class Test {
    public static void main(String[] args) {
        // 将luckysheet json文件转为excel
        LuckysheetConverter.luckysheetToExcel("/path/luckysheet.json", "/path/excel.xlsx");
        
        // 将excel转为luckysheet json文件 
        LuckysheetConverter.excelToLuckySheetFile("/path/excel.xlsx", "/path/luckysheet.json");
    }
}
```