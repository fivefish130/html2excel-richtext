# HTML to Excel Rich Text Examples

This module contains complete working examples of how to use the html2excel-richtext library.

## Examples Included

### 1. CoreExample.java
Demonstrates core functionality:
- ✅ Basic HTML conversion (bold, italic, underline)
- ✅ Rich text with colors, fonts, and sizes
- ✅ Lists (ordered and unordered)
- ✅ Tables
- ✅ Hyperlinks
- ✅ Background colors
- ✅ Images with async download

**Run it:**
```bash
mvn exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.CoreExample"
```

**Output:** 7 Excel files demonstrating different features

### 2. JxlsExample.java
Shows JXLS template integration:
- ✅ Using `jx:html()` command in templates
- ✅ Processing templates with HTML fields
- ✅ Custom configuration options
- ✅ Working with lists of objects

**Note:** Requires creating a JXLS template first. See code comments for instructions.

**Run it:**
```bash
mvn exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.JxlsExample"
```

### 3. EasyExcelExample.java
Demonstrates EasyExcel annotation integration:
- ✅ `@HtmlCell` annotation usage
- ✅ Multiple HTML fields
- ✅ Image download configuration
- ✅ Custom truncation settings

**Run it:**
```bash
mvn exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.EasyExcelExample"
```

**Output:** 3 Excel files with different annotation configurations

## Quick Start

### Option 1: Run all Core examples
```bash
cd html2excel-richtext-examples
mvn clean compile exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.CoreExample"
```

This will create 7 Excel files:
- `example1-basic.xlsx` - Basic formatting
- `example2-richtext.xlsx` - Colors, fonts, sizes
- `example3-lists.xlsx` - Ordered and unordered lists
- `example4-tables.xlsx` - Table formatting
- `example5-hyperlinks.xlsx` - Clickable links
- `example6-backgrounds.xlsx` - Cell background colors
- `example7-images.xlsx` - Embedded images (requires internet)

### Option 2: Run EasyExcel examples
```bash
mvn clean compile exec:java -Dexec.mainClass="io.github.fivefish130.html2excel.richtext.examples.EasyExcelExample"
```

This will create 3 Excel files demonstrating different annotation usages.

### Option 3: Copy example code
All example classes are fully self-contained and can be copied directly into your project.

## Building Examples

```bash
# Build only examples module
cd html2excel-richtext-examples
mvn clean compile

# Build entire project including examples
cd ..
mvn clean install
```

## Code Structure

```
src/main/java/io/github/fivefish130/html2excel/richtext/examples/
├── CoreExample.java           # Basic usage examples
├── JxlsExample.java          # JXLS template examples
└── EasyExcelExample.java     # EasyExcel annotation examples
```

## Dependencies

All dependencies are already configured in the `pom.xml`:
- html2excel-richtext-core
- html2excel-richtext-jxls
- html2excel-richtext-easyexcel

## Tips

1. **Internet Connection**: Image download examples require internet connection
2. **File Paths**: Examples create output files in the current directory
3. **Modify Examples**: Feel free to modify the HTML content to test different scenarios
4. **JXLS Templates**: For JXLS examples, you need to create template files first

## Common HTML Tags Supported

| Feature | Tag/CSS | Example |
|---------|---------|---------|
| Bold | `<b>`, `<strong>` | `<b>Bold</b>` |
| Italic | `<i>`, `<em>` | `<i>Italic</i>` |
| Underline | `<u>` | `<u>Underline</u>` |
| Color | `style="color:..."` | `<span style='color:red'>Red</span>` |
| Font | `style="font-family:..."` | `<span style='font-family:Arial'>Arial</span>` |
| Size | `style="font-size:..."` | `<span style='font-size:14px'>14px</span>` |
| Background | `style="background-color:..."` | `<p style='background-color:#FFFF00'>Highlighted</p>` |
| Link | `<a href="...">` | `<a href='url'>Link</a>` |
| Image | `<img src="...">` | `<img src='url'/>` |
| Break | `<br>`, `<p>` | `Line1<br/>Line2` |
| List | `<ul>`, `<ol>`, `<li>` | `<ul><li>Item</li></ul>` |
| Table | `<table>`, `<tr>`, `<td>` | `<table><tr><td>Cell</td></tr></table>` |

## Need Help?

- See main README: [../../README.md](../../README.md)
- Open an issue: [GitHub Issues](https://github.com/fivefish130/html2excel-richtext/issues)
- Read the docs: [Documentation](https://github.com/fivefish130/html2excel-richtext)
