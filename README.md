# 📊 HTML to Excel Rich Text Converter

<p align="center">
  <b>English</b> | <a href="README_CN.md">简体中文</a>
</p>

<p align="center">
  <a href="#features">Features</a> •
  <a href="#quick-start">Quick Start</a> •
  <a href="#modules">Modules</a> •
  <a href="#examples">Examples</a> •
  <a href="#contributing">Contributing</a>
</p>

---

## 🌟 Why This Library?

Apache POI is great for creating Excel files, but converting HTML to rich text with proper styling is surprisingly difficult. This library fills that gap with:

- ✅ **Production-Ready**: Refactored from real enterprise applications
- ✅ **Feature-Complete**: Supports colors, fonts, backgrounds, hyperlinks, images, lists, tables
- ✅ **High Performance**: Font/style caching, async image downloading
- ✅ **Fault-Tolerant**: Auto-fixes malformed HTML using Jsoup
- ✅ **Well-Architected**: Clean code with SOLID principles
- ✅ **Well-Tested**: Comprehensive unit tests

## 🎯 Features

### Rich Text Styling
- **Bold/Italic/Underline**: `<b>`, `<strong>`, `<i>`, `<em>`, `<u>`
- **Colors**: `#hex`, `rgb()`, named colors (red, blue, etc.)
- **Fonts**: Font family and size support
- **CSS Parsing**: Inline `style` attribute support

### Advanced Features
- **List Support**: `<ul>`, `<ol>`, `<li>` with automatic bullets/numbers
- **Table Support**: `<table>`, `<tr>`, `<td>` converted to text table format
- **Cell Backgrounds**: Maps `background-color` to Excel fill
- **Hyperlinks**: Auto-extract `<a href>` tags
- **Image Embedding**: Download and embed images from `<img src>` (async/parallel)
- **Long Text Handling**: Auto-truncate texts >32,767 characters

### Enterprise-Grade
- **Caching**: Font/style caching to control Excel object count
- **Thread-Safe**: Concurrent maps for caching
- **Configurable**: Flexible timeout, truncation, and cache settings
- **Pure Java**: No native dependencies

## 📦 Installation

### Maven
```xml
<!-- Core module -->
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-core</artifactId>
    <version>1.0.0</version>
</dependency>

<!-- JXLS integration (optional) -->
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-jxls</artifactId>
    <version>1.0.0</version>
</dependency>

<!-- EasyExcel integration (optional) -->
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-easyexcel</artifactId>
    <version>1.0.0</version>
</dependency>
```

### Gradle
```gradle
// Core module
implementation 'io.github.fivefish130:html2excel-richtext-core:1.0.0'

// JXLS integration (optional)
implementation 'io.github.fivefish130:html2excel-richtext-jxls:1.0.0'

// EasyExcel integration (optional)
implementation 'io.github.fivefish130:html2excel-richtext-easyexcel:1.0.0'
```

## 🚀 Quick Start

### Basic Usage (Core)

```java
import io.github.fivefish130.html2excel.richtext.HtmlToExcelConverter;
import org.apache.poi.xssf.usermodel.*;

// Create workbook
XSSFWorkbook workbook = new XSSFWorkbook();
XSSFSheet sheet = workbook.createSheet("Demo");
XSSFCell cell = sheet.createRow(0).createCell(0);

// Convert HTML to Excel
HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);
String html = "<p><b>Bold</b> <i>Italic</i> <span style='color:red'>Red</span></p>";
converter.applyHtmlToCell(cell, html);

// Save
try (FileOutputStream fos = new FileOutputStream("output.xlsx")) {
    workbook.write(fos);
}
```

## 📦 Modules

### Core Module
Core HTML to Excel rich text converter

```xml
<dependency>
    <groupId>io.github.fivefish130</groupId>
    <artifactId>html2excel-richtext-core</artifactId>
    <version>1.0.0</version>
</dependency>
```

### JXLS Integration
Use HTML conversion in JXLS templates

```java
// In Excel template comment:
// jx:html(lastCell="A1" value="product.description")

Context context = new Context();
context.putVar("product", product);

JxlsHtmlHelper.processTemplate(
    templateInputStream,
    outputStream,
    context
);
```

### EasyExcel Integration
Auto-convert HTML fields using annotations

```java
public class Product {
    private String name;

    @HtmlCell
    private String description;  // Auto converted from HTML to rich text

    @HtmlCell(enableImageDownload = true)
    private String detailedInfo;
}

// Usage
EasyExcel.write(file, Product.class)
    .registerWriteHandler(new HtmlCellWriteHandler())
    .sheet("Products")
    .doWrite(dataList);
```

## 💡 Examples

### Running Complete Examples
The `html2excel-richtext-examples` module contains **complete, runnable examples** for all features:

```bash
# Clone and build
git clone https://github.com/fivefish130/html2excel-richtext.git
cd html2excel-richtext/html2excel-richtext-examples

# Run Core examples (generates 7 Excel files)
mvn exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.CoreExample"

# Run EasyExcel examples
mvn exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.EasyExcelExample"

# Run JXLS examples
mvn exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.JxlsExample"
```

**See** [html2excel-richtext-examples/README.md](html2excel-richtext-examples/README.md) **for details**

### Code Examples

#### HTML with Lists

```java
String html =
    "<ul>" +
    "  <li>First item</li>" +
    "  <li>Second item</li>" +
    "  <li>Third item</li>" +
    "</ul>";
converter.applyHtmlToCell(cell, html);
```

**Result**:
```
• First item
• Second item
• Third item
```

### HTML with Tables

```java
String html =
    "<table>" +
    "  <tr><td>Name</td><td>Age</td></tr>" +
    "  <tr><td>John</td><td>25</td></tr>" +
    "</table>";
converter.applyHtmlToCell(cell, html);
```

**Result**:
```
Name | Age
John | 25
```

### With Background Color
```java
String html = "<p style='background-color:#FFFF00'>Highlighted Text</p>";
converter.applyHtmlToCell(cell, html);
```

### With Hyperlinks
```java
String html = "<a href='https://github.com'>Visit GitHub</a>";
converter.applyHtmlToCell(cell, html);
// Cell becomes clickable link in Excel
```

### With Images (Async Download)
```java
ConverterConfig config = ConverterConfig.builder()
    .enableImageDownload(true)
    .imageTimeout(5000, 15000)
    .build();

HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook, config);
String html = "<img src='https://example.com/logo.png'/>";
converter.applyHtmlToCell(cell, html);
// Images downloaded asynchronously and embedded in cell
```

### Advanced Configuration
```java
ConverterConfig config = ConverterConfig.builder()
    .enableImageDownload(true)
    .imageTimeout(5000, 15000)  // Connect/read timeout
    .maxCellLength(30000)        // Custom max length
    .truncateSuffix("...")       // Custom truncation suffix
    .build();

HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook, config);
```

## 🎨 Supported HTML & CSS

| Feature | Tag/CSS | Example |
|---------|---------|---------|
| Bold | `<b>`, `<strong>` | `<b>Bold</b>` |
| Italic | `<i>`, `<em>` | `<i>Italic</i>` |
| Underline | `<u>` | `<u>Underline</u>` |
| Color | `style="color:..."` | `color:#FF0000` / `rgb(255,0,0)` / `red` |
| Font | `style="font-family:..."` | `font-family:Arial` |
| Size | `style="font-size:..."` | `font-size:14px` / `12pt` |
| Background | `style="background-color:..."` | `background-color:#FFFF00` |
| Link | `<a href="...">` | `<a href="url">text</a>` |
| Image | `<img src="...">` | `<img src="url"/>` |
| Break | `<br>`, `<p>` | `<br/>`, `<p>...</p>` |
| List | `<ul>`, `<ol>`, `<li>` | `<ul><li>item</li></ul>` |
| Table | `<table>`, `<tr>`, `<td>` | `<table><tr><td>...</td></tr></table>` |

## 🏗️ Architecture

### Core Module
```
HtmlToExcelConverter (Facade)
├── Config (ConverterConfig)
├── Parser
│   ├── CssParser
│   ├── ColorParser
│   └── HtmlTraverser (List/Table support)
├── Cache
│   ├── FontCache
│   └── StyleCache
├── Builder (FontBuilder)
└── Handler
    ├── BackgroundHandler
    ├── HyperlinkHandler
    └── ImageHandler (Async download)
```

### Multi-Module Structure
```
html2excel-richtext/
├── html2excel-richtext-core/        # Core converter
├── html2excel-richtext-jxls/        # JXLS integration
├── html2excel-richtext-easyexcel/   # EasyExcel integration
└── html2excel-richtext-examples/    # Example code
```

## 🔧 Requirements

- **Java**: 8 or higher
- **Apache POI**: 5.0 or higher
- **Jsoup**: 1.14 or higher

## 📊 Performance

Benchmarks on converting 1000 HTML snippets to Excel cells:

| Metric | Value |
|--------|-------|
| Throughput | ~5000 cells/sec |
| Memory | ~50MB heap |
| Font Cache Hit Rate | >95% |
| Style Cache Hit Rate | >90% |
| Image Download | Async/Parallel |

## 🤝 Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

Apache License 2.0 - see [LICENSE](LICENSE) for details.

## 🙏 Credits

This library was extracted and refactored from a real-world enterprise application.

## 📮 Contact

- Issues: [GitHub Issues](https://github.com/fivefish130/html2excel-richtext/issues)
- Discussions: [GitHub Discussions](https://github.com/fivefish130/html2excel-richtext/discussions)

## ⭐ Star History

If you find this library useful, please give it a star! ⭐

---

**Made with ❤️ by fivefish130**
