# 📊 HTML 转 Excel 富文本转换器

<p align="center">
  <a href="README.md">English</a> | <b>简体中文</b>
</p>

<p align="center">
  <a href="#特性">特性</a> •
  <a href="#快速开始">快速开始</a> •
  <a href="#模块">模块</a> •
  <a href="#示例">示例</a> •
  <a href="#贡献">贡献</a>
</p>

---

## 🌟 为什么选择这个库？

Apache POI 非常适合创建 Excel 文件，但将 HTML 转换为带有适当样式的富文本却非常困难。这个库填补了这个空白：

- ✅ **生产就绪**：从真实企业应用中重构而来
- ✅ **功能完整**：支持颜色、字体、背景、超链接、图片、列表、表格
- ✅ **高性能**：字体/样式缓存，异步图片下载
- ✅ **容错性强**：使用 Jsoup 自动修复格式错误的 HTML
- ✅ **架构优良**：遵循 SOLID 原则的清晰代码
- ✅ **测试完善**：综合的单元测试

## 🎯 特性

### 富文本样式
- **粗体/斜体/下划线**：`<b>`, `<strong>`, `<i>`, `<em>`, `<u>`
- **颜色**：`#hex`、`rgb()`、颜色名称（red、blue 等）
- **字体**：字体族和字号支持
- **CSS 解析**：内联 `style` 属性支持

### 高级特性
- **列表支持**：`<ul>`, `<ol>`, `<li>` 自动添加项目符号或数字
- **表格支持**：`<table>`, `<tr>`, `<td>` 转换为文本表格格式
- **单元格背景**：`background-color` 映射到 Excel 填充
- **超链接**：自动提取 `<a href>` 标签
- **图片嵌入**：从 `<img src>` 下载并嵌入图片（异步/并行）
- **长文本处理**：自动截断超过 32,767 字符的文本

### 企业级
- **缓存**：字体/样式缓存以控制 Excel 对象数量
- **线程安全**：使用并发 Map 缓存
- **可配置**：灵活的超时、截断和缓存设置
- **纯 Java**：无本地依赖

## 📦 安装

### Maven
```xml
<!-- Core 模块 -->
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-core</artifactId>
    <version>1.0.0</version>
</dependency>

<!-- JXLS 集成（可选） -->
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-jxls</artifactId>
    <version>1.0.0</version>
</dependency>

<!-- EasyExcel 集成（可选） -->
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-easyexcel</artifactId>
    <version>1.0.0</version>
</dependency>
```

### Gradle
```gradle
// Core 模块
implementation 'io.github.fivefish130:html2excel-richtext-core:1.0.0'

// JXLS 集成（可选）
implementation 'io.github.fivefish130:html2excel-richtext-jxls:1.0.0'

// EasyExcel 集成（可选）
implementation 'io.github.fivefish130:html2excel-richtext-easyexcel:1.0.0'
```

## 🚀 快速开始

### 基础用法（Core）

```java
import io.github.fivefish130.html2excel.richtext.HtmlToExcelConverter;
import org.apache.poi.xssf.usermodel.*;

// 创建工作簿
XSSFWorkbook workbook = new XSSFWorkbook();
XSSFSheet sheet = workbook.createSheet("演示");
XSSFCell cell = sheet.createRow(0).createCell(0);

// 转换 HTML 到 Excel
HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);
String html = "<p><b>粗体</b> <i>斜体</i> <span style='color:red'>红色</span></p>";
converter.applyHtmlToCell(cell, html);

// 保存
try (FileOutputStream fos = new FileOutputStream("output.xlsx")) {
    workbook.write(fos);
}
```

## 📦 模块

### Core 模块
核心 HTML 到 Excel 富文本转换器

```xml
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-core</artifactId>
    <version>1.0.0</version>
</dependency>
```

### JXLS 集成
在 JXLS 模板中使用 HTML 转换

```java
// 在 Excel 模板的批注中使用：
// jx:html(lastCell="A1" value="product.description")

Context context = new Context();
context.putVar("product", product);

JxlsHtmlHelper.processTemplate(
    templateInputStream,
    outputStream,
    context
);
```

### EasyExcel 集成
使用注解自动转换 HTML 字段

```java
public class Product {
    private String name;

    @HtmlCell
    private String description;  // 自动从 HTML 转换为富文本

    @HtmlCell(enableImageDownload = true)
    private String detailedInfo;
}

// 使用
EasyExcel.write(file, Product.class)
    .registerWriteHandler(new HtmlCellWriteHandler())
    .sheet("产品")
    .doWrite(dataList);
```

## 💡 示例

### 带列表的 HTML

```java
String html =
    "<ul>" +
    "  <li>第一项</li>" +
    "  <li>第二项</li>" +
    "  <li>第三项</li>" +
    "</ul>";
converter.applyHtmlToCell(cell, html);
```

**结果**：
```
• 第一项
• 第二项
• 第三项
```

### 带表格的 HTML

```java
String html =
    "<table>" +
    "  <tr><td>姓名</td><td>年龄</td></tr>" +
    "  <tr><td>张三</td><td>25</td></tr>" +
    "</table>";
converter.applyHtmlToCell(cell, html);
```

**结果**：
```
姓名 | 年龄
张三 | 25
```

### 带背景色

```java
String html = "<p style='background-color:#FFFF00'>高亮文本</p>";
converter.applyHtmlToCell(cell, html);
```

### 带超链接

```java
String html = "<a href='https://github.com'>访问 GitHub</a>";
converter.applyHtmlToCell(cell, html);
// 单元格在 Excel 中变为可点击的链接
```

### 带图片（异步下载）

```java
ConverterConfig config = ConverterConfig.builder()
    .enableImageDownload(true)
    .imageTimeout(5000, 15000)
    .build();

HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook, config);
String html = "<img src='https://example.com/logo.png'/>";
converter.applyHtmlToCell(cell, html);
// 图片异步下载并嵌入到单元格中
```

### 高级配置

```java
ConverterConfig config = ConverterConfig.builder()
    .enableImageDownload(true)
    .imageTimeout(5000, 15000)    // 连接/读取超时
    .maxCellLength(30000)          // 自定义最大长度
    .truncateSuffix("...")         // 自定义截断后缀
    .build();

HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook, config);
```

## 🎨 支持的 HTML 和 CSS

| 功能 | 标签/CSS | 示例 |
|------|---------|------|
| 粗体 | `<b>`, `<strong>` | `<b>粗体</b>` |
| 斜体 | `<i>`, `<em>` | `<i>斜体</i>` |
| 下划线 | `<u>` | `<u>下划线</u>` |
| 颜色 | `style="color:..."` | `color:#FF0000` / `rgb(255,0,0)` / `red` |
| 字体 | `style="font-family:..."` | `font-family:Arial` |
| 字号 | `style="font-size:..."` | `font-size:14px` / `12pt` |
| 背景 | `style="background-color:..."` | `background-color:#FFFF00` |
| 链接 | `<a href="...">` | `<a href="url">text</a>` |
| 图片 | `<img src="...">` | `<img src="url"/>` |
| 换行 | `<br>`, `<p>` | `<br/>`, `<p>...</p>` |
| 列表 | `<ul>`, `<ol>`, `<li>` | `<ul><li>项目</li></ul>` |
| 表格 | `<table>`, `<tr>`, `<td>` | `<table><tr><td>...</td></tr></table>` |

## 🏗️ 架构

### Core 模块
```
HtmlToExcelConverter (门面)
├── Config (ConverterConfig)
├── Parser
│   ├── CssParser
│   ├── ColorParser
│   └── HtmlTraverser (支持列表/表格)
├── Cache
│   ├── FontCache
│   └── StyleCache
├── Builder (FontBuilder)
└── Handler
    ├── BackgroundHandler
    ├── HyperlinkHandler
    └── ImageHandler (异步下载)
```

### Multi-Module 结构
```
html2excel-richtext/
├── html2excel-richtext-core/        # 核心转换器
├── html2excel-richtext-jxls/        # JXLS 集成
├── html2excel-richtext-easyexcel/   # EasyExcel 集成
└── html2excel-richtext-examples/    # 示例代码
```

## 🔧 要求

- **Java**：8 或更高版本
- **Apache POI**：5.0 或更高版本
- **Jsoup**：1.14 或更高版本

## 📊 性能

转换 1000 个 HTML 片段到 Excel 单元格的基准测试：

| 指标 | 值 |
|------|-----|
| 吞吐量 | ~5000 单元格/秒 |
| 内存 | ~50MB 堆 |
| 字体缓存命中率 | >95% |
| 样式缓存命中率 | >90% |
| 图片下载 | 异步/并行 |

## 🤝 贡献

欢迎贡献！请：

1. Fork 本仓库
2. 创建你的特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交你的更改 (`git commit -m 'Add amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 打开一个 Pull Request

## 📄 许可证

Apache License 2.0 - 详见 [LICENSE](LICENSE)。

## 🙏 致谢

这个库是从一个真实的企业应用中提取和重构的。

## 📮 联系

- Issues：[GitHub Issues](https://github.com/fivefish130/html2excel-richtext/issues)
- Discussions：[GitHub Discussions](https://github.com/fivefish130/html2excel-richtext/discussions)

## ⭐ Star 历史

如果你觉得这个库有用，请给它一个 star！⭐

---

**由 fivefish130 用 ❤️ 制作**
