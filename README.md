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
  - sheet标签颜色
  - 行隐藏
  - 列隐藏
  - 行冻结
  - 列冻结
  - 网格线显示/隐藏
  - 默认行高/列宽
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
  - 缩小字体填充
- 公式
  - 绝大多数公式，少量公式存在luckysheet和excel兼容问题，大家有发现可以反馈
- 批注
  - 批注内容
  - 批注锚点
- 图片
  - 常见的各种格式图片
  - 图片位置映射
- 超链接
  - 外部链接（URL、邮件）
  - 工作表内部链接
  - 链接提示文字
- 数据验证
  - 下拉列表
  - 复选框
  - 数字范围
  - 文本内容
  - 文本长度
  - 日期
  - 身份证/手机号（有损，公式拟合）
  - 输入提示和错误提示
- 条件格式
  - 单元格颜色规则（大于、小于、介于、等于等）
  - 数据条
  - 色阶（2色/3色）
  - 图标集
- 自动筛选
  - 筛选范围（列条件有损）
- 工作表保护
  - 密码保护
  - 各类操作权限（选择、格式、插入、删除、排序、筛选等）
- 行列分组（大纲）
  - 多级分组
  - 折叠/展开状态
- 命名范围
  - 工作簿级命名范围
- 图表（有损）
  - 柱状图、折线图、饼图基础结构
- 数据透视表（有损）
  - 行/列/值字段基础结构
- 迷你图（有损）
  - 折线/柱状/盈亏类型