# CLAUDE.md

本文件为 Claude Code (claude.ai/code) 在此代码库中工作时提供指导。

## 项目概述

luckysheet-lib 是一个用于在 Excel (xlsx) 文件和 Luckysheet JSON 格式之间进行转换的 Java 库。这是一个双向转换器，可以保留电子表格的数据、样式、公式、图片和批注。

**核心功能：**
- Excel → Luckysheet JSON（用于基于 Web 的电子表格编辑）
- Luckysheet JSON → Excel（用于下载/导出）

**已发布到 Maven Central：** `io.github.autoffice:luckysheet-lib`

## 构建命令

本项目使用 Maven 包装器（`mvnw`）。在 Unix 系统上使用 `./mvnw`，在 Windows 上使用 `mvnw.cmd`。

### 基本命令

```bash
# 运行所有测试
./mvnw test

# 使用 Java 11+ 运行测试（新版 JDK 必需）
./mvnw test -DargLine="--add-opens=java.base/java.lang=ALL-UNNAMED --add-opens=java.base/java.util=ALL-UNNAMED --add-opens=java.base/sun.reflect.annotation=ALL-UNNAMED"

# 构建并安装到本地 Maven 仓库
./mvnw install

# 运行特定测试类
./mvnw test -Dtest=LuckysheetConverterTest

# 运行特定测试方法
./mvnw test -Dtest=LuckysheetConverterTest#fullTesting

# 生成 Javadoc
./mvnw javadoc:javadoc

# 检查代码风格（PMD）
./mvnw pmd:check

# 检查许可证头
./mvnw license:check

# 生成代码覆盖率报告（JaCoCo）
./mvnw test jacoco:report
# 报告位置：target/site/jacoco/index.html
```

### 构建阶段

构建过程包含多个自动运行的验证步骤：
- **validate 阶段：** PMD 代码风格检查、许可证头验证
- **test 阶段：** JUnit 测试和 JaCoCo 覆盖率
- **compile 阶段：** 源码 JAR 生成

## 架构设计

### 转换流程

**Excel → Luckysheet：**
```
XSSFWorkbook (POI) 
  → WorkbookMapperToLuckySheet 
    → SheetMapperToLuckySheet (每个工作表)
      → CellMapperToLuckySheet (每个单元格)
      → ImageMapperToLuckySheet (图片)
  → LuckyFile (模型)
```

**Luckysheet → Excel：**
```
LuckyFile (模型)
  → WorkbookMapperToExcel
    → SheetMapperToExcel (每个工作表)
      → CellMapperToExcel (每个单元格)
      → ImageMapperToExcel (图片)
  → XSSFWorkbook (POI)
```

### 核心组件

**入口点：**
- `LuckysheetConverter` - 公共 API，包含所有转换操作的静态方法

**映射器（双向转换逻辑）：**
- `WorkbookMapperToLuckySheet` / `WorkbookMapperToExcel` - 工作簿级别转换
- `SheetMapperToLuckySheet` / `SheetMapperToExcel` - 工作表级别转换（行、列、冻结窗格、合并单元格）
- `CellMapperToLuckySheet` / `CellMapperToExcel` - 单元格级别转换（数据、样式、公式、批注）
- `ImageMapperToLuckySheet` / `ImageMapperToExcel` - 图片处理和定位

**工厂类：**
- `PoiFactory` - 创建 POI 对象（XSSFWorkbook、XSSFSheet 等）
- `LuckySheetFactory` - 创建 Luckysheet 模型对象

**模型（`io.github.autoffice.luckysheet.model`）：**
- `LuckyFile` - 顶层容器（相当于工作簿）
- `LuckySheet` - 工作表数据和配置
- `CellData` - 单元格内容和样式
- `SheetImage` - 图片元数据和位置
- 各种枚举类型：单元格类型、对齐方式、边框等

**工具类：**
- `PoiUtil` - POI 特定的辅助方法
- `NumberUtil` - 数字格式化和转换
- `DateUtil` - 日期/时间处理
- `ImageUtil` - 图片处理
- `JSONUtil` - JSON 序列化

### 设计模式

1. **映射器模式：** 为每个转换方向和层级（工作簿/工作表/单元格）使用独立的映射器类
2. **工厂模式：** 集中创建 POI 和 Luckysheet 模型对象
3. **模型驱动：** `io.github.autoffice.luckysheet.model` 包中的丰富领域模型表示 Luckysheet 结构

## 开发说明

### Java 版本
- **目标版本：** Java 8（最大兼容性）
- **测试环境：** CI 在 Java 8、11、17 和 21 上测试
- **本地开发：** Java 11+ 需要额外的 JVM 参数（见上面的测试命令）

### 依赖项
- **Apache POI 5.4.0** - Excel 文件操作
- **Jackson** - JSON 序列化
- **Lombok** - 减少样板代码（仅编译时）
- **JUnit 5** - 测试框架

### 代码质量
- PMD 强制执行代码风格规则
- 所有 `.java` 文件必须包含 Apache 2.0 许可证头
- JaCoCo 跟踪测试覆盖率（报告上传到 Codecov）

### 测试
- 测试资源位于 `src/test/resources/`（示例 Excel 和 JSON 文件）
- 测试输出到 `./target/` 目录
- 主要测试：`LuckysheetConverterTest#fullTesting` - 往返转换测试

### 常见陷阱
- **富文本：** Luckysheet 使用内联文本格式；POI 使用 XSSFRichTextString
- **颜色：** Luckysheet 使用十六进制字符串；POI 使用带 ARGB 的 XSSFColor
- **公式：** Luckysheet 和 Excel 之间存在一些公式兼容性问题
- **日期格式：** 需要仔细映射 POI 日期格式和 Luckysheet 格式代码
