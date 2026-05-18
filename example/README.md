# Luckysheet-lib 导入导出示例

本示例演示如何使用 `luckysheet-lib` 实现 Excel 与 Luckysheet JSON 格式的双向转换，包含完整的前后端实现。

## 前置要求

- JDK 8+
- Maven 3.6+
- Node.js 16+

## 快速开始

### 1. 安装 luckysheet-lib 到本地 Maven 仓库

```bash
cd ../
mvn clean install -DskipTests
```

### 2. 启动后端

```bash
cd example/backend
mvn clean install
mvn spring-boot:run
```

后端将在 `http://localhost:8181` 启动。

### 3. 启动前端

```bash
cd example/frontend
npm install
npm run dev
```

前端将在 `http://localhost:5173` 启动，浏览器打开后自动加载示例数据。

## 项目结构

```
example/
├── backend/                                    # Spring Boot 2.7 后端
│   ├── pom.xml
│   └── src/main/
│       ├── java/com/example/luckysheet/
│       │   ├── LuckysheetApplication.java
│       │   ├── config/CorsConfig.java
│       │   ├── controller/LuckysheetController.java
│       │   └── service/LuckysheetService.java
│       └── resources/
│           ├── application.yml
│           └── full.json                       # 示例数据
├── frontend/                                   # Vue 3 + Vite
│   ├── package.json
│   ├── vite.config.js
│   ├── index.html
│   └── src/
│       ├── main.js
│       └── App.vue
└── README.md
```

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/excel/upload` | 上传 Excel 文件，返回 Luckysheet JSON |
| POST | `/api/excel/export` | 接收 Luckysheet JSON，返回 Excel 文件 |
| GET  | `/api/excel/sample` | 获取示例 Luckysheet JSON 数据 |
| GET  | `/api/excel/health` | 健康检查 |

### curl 示例

```bash
# 上传 Excel
curl -X POST http://localhost:8181/api/excel/upload \
  -F "file=@test.xlsx" -o output.json

# 导出 Excel
curl -X POST http://localhost:8181/api/excel/export \
  -H "Content-Type: application/json" \
  -d @luckysheet.json --output result.xlsx

# 获取示例数据
curl http://localhost:8181/api/excel/sample
```

## 支持的特性

### 基础特性
- 单元格数据（文本、数字、公式）
- 单元格样式（字体、颜色、边框、对齐）
- 合并单元格、行高列宽、冻结行列、多工作表

### 高级特性
- 超链接 - 外部链接和工作表内链接
- 数据验证 - 下拉列表、数字范围、日期、文本长度
- 条件格式 - 颜色规则、数据条、色阶、图标集
- 自动筛选 - 筛选范围保留
- 工作表保护 - 锁定单元格、限制编辑权限
- 行列分组 - 多级分组和折叠
- 命名范围 - 工作簿级别的命名区域

### 图表和数据分析（有损转换）
- 图表 - 基础柱状图、折线图、饼图
- 数据透视表 - 基础结构保留
- 迷你图 - 通过 Excel 扩展保存

## 技术栈

| 层 | 技术 |
|----|------|
| 后端 | Java 8+, Spring Boot 2.7, luckysheet-lib 1.2.0-SNAPSHOT, Apache POI 5.4.0 |
| 前端 | Vue 3, Vite 5, Luckysheet (CDN), Axios |

## 配置说明

### 后端 (application.yml)

```yaml
server:
  port: 8181

spring:
  servlet:
    multipart:
      max-file-size: 10MB
      max-request-size: 10MB
```

### 前端 (vite.config.js)

开发模式下通过 Vite proxy 将 `/api` 请求转发到后端 `http://localhost:8181`。

## 常见问题

**后端启动失败：找不到 luckysheet-lib**

先在项目根目录执行 `mvn clean install -DskipTests` 安装到本地仓库。

**前端无法连接后端**

确认后端在 8181 端口运行：`curl http://localhost:8181/api/excel/health`

**上传大文件失败**

修改 `application.yml` 中 `max-file-size` 和 `max-request-size`。

**处理大型 Excel 文件时 OOM**

增加 JVM 堆内存：`java -Xmx2g -jar app.jar`

## 许可证

Apache License 2.0
