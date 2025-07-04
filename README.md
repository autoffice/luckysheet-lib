luckysheet-lib
======================
[![Java CI](https://github.com/autoffice/luckysheet-lib/actions/workflows/ci.yml/badge.svg)](https://github.com/autoffice/luckysheet-lib/actions/workflows/ci.yml)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)
[![codecov](https://codecov.io/gh/autoffice/luckysheet-lib/graph/badge.svg?token=DP021D9LUG)](https://codecov.io/gh/autoffice/luckysheet-lib)
[![CodeQL](https://github.com/autoffice/luckysheet-lib/actions/workflows/github-code-scanning/codeql/badge.svg)](https://github.com/autoffice/luckysheet-lib/actions/workflows/github-code-scanning/codeql)
[![Maven Central](https://img.shields.io/maven-central/v/io.github.autoffice/luckysheet-lib.svg?label=Maven%20Central)](https://search.maven.org/artifact/io.github.autoffice/luckysheet-lib)

luckysheet-lib是Luckysheet的Java库，支持包括excel xlsx导入luckysheet和luckysheet导出为xlsx excel文件。

**如果你觉的还不错，欢迎star 或者 fork**

## 使用方法
pom.xml引入luckysheet-lib依赖, [![最新版本](https://img.shields.io/maven-central/v/io.github.autoffice/luckysheet-lib.svg?label=%E6%9C%80%E6%96%B0%E7%89%88%E6%9C%AC)](https://search.maven.org/artifact/io.github.autoffice/luckysheet-lib)

```xml
            <dependency>
                <groupId>io.github.autoffice</groupId>
                <artifactId>luckysheet-lib</artifactId>
                <version>最新版本</version>
            </dependency>
```

使用`LuckysheetConverter`类中对应的导入、导出方法即可，多种方法总有一种适合你，例如：
```java
import io.github.autoffice.luckysheet.LuckysheetConverter;

public class Test {
  public static void main(String[] args) throws IOException, InvalidFormatException {
    // 将luckysheet json文件转为excel
    LuckysheetConverter.luckysheetToExcel("/path/luckysheet.json", "/path/excel.xlsx");

    // 将luckysheet json文件转为OutputStream
    LuckysheetConverter.luckysheetToExcel("/path/luckysheet.json", Files.newOutputStream(Paths.get("/path/excel.xlsx")));

    // 将luckysheet json文件转为luckysheet对象
    LuckyFile luckyFile = LuckysheetConverter.readAsLuckyFile("/path/luckysheet.json");

    // 将excel转为luckysheet json文件
    LuckysheetConverter.excelToLuckySheetFile("/path/excel.xlsx", "/path/luckysheet.json");

    // 将excel文件转为luckysheet对象
    LuckyFile luckyFile1 = LuckysheetConverter.excelToLuckySheet("/path/excel.xlsx");

    // 将excel文件转为luckysheet json
    String json = LuckysheetConverter.excelToLuckySheetJson("/path/excel.xlsx");
  }
}
```

## 支持功能列表
当前已经完成以下表格特性的导入和导出支持:
- 导入导出
  - xlsx文件导入为luckysheet json
  - luckysheet json下载为xlsx文件
- sheet数据和样式
  - 多sheet
  - sheet名称
  - 行隐藏
  - 列隐藏
  - 行冻结
  - 列冻结
- 单元格数据和样式
  - 单元格数据
  - 单元格背景颜色
  - 边框颜色
  - 边框样式
  - 字体
  - 字体颜色
  - 富文本文字
  - 加粗
  - 斜体
  - 下划线
  - 删除线
  - 单元格合并
  - 数字格式
  - 日期格式
  - 各种方向文本对齐
  - 自动换行
  - 文字旋转
- 公式
  -  绝大多数公式，少量公式存证luckysheet和excel兼容问题，大家有发现可以反馈
- 批注
  - 批注内容
  - 批注锚点
- 图片
  - 常见的各种格式图片
  - 图片位置映射