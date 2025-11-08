package io.github.fivefish130.html2excel.richtext.examples;

import io.github.fivefish130.html2excel.richtext.HtmlToExcelConverter;
import io.github.fivefish130.html2excel.richtext.config.ConverterConfig;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;

/**
 * Basic examples of using the Core module
 *
 * @author fivefish130
 */
public class CoreExample {

    public static void main(String[] args) throws Exception {
        // Example 1: Basic HTML conversion
        basicExample();

        // Example 2: Rich text with multiple styles
        richTextExample();

        // Example 3: Lists
        listExample();

        // Example 4: Tables
        tableExample();

        // Example 5: Hyperlinks
        hyperlinkExample();

        // Example 6: Background colors
        backgroundExample();

        // Example 7: Images (async download)
        imageExample();

        System.out.println("All examples completed! Check the output files.");
    }

    /**
     * Example 1: Basic HTML to Excel conversion
     */
    public static void basicExample() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Basic Example");
        XSSFCell cell = sheet.createRow(0).createCell(0);

        HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);
        String html = "<p><b>Bold</b> <i>Italic</i> <u>Underline</u></p>";
        converter.applyHtmlToCell(cell, html);

        sheet.setColumnWidth(0, 8000);

        try (FileOutputStream fos = new FileOutputStream("example1-basic.xlsx")) {
            workbook.write(fos);
        }
        workbook.close();

        System.out.println("✓ Example 1: Basic conversion -> example1-basic.xlsx");
    }

    /**
     * Example 2: Rich text with colors, fonts, and sizes
     */
    public static void richTextExample() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Rich Text");

        int rowNum = 0;

        // Color examples
        XSSFCell cell1 = sheet.createRow(rowNum++).createCell(0);
        HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);
        converter.applyHtmlToCell(cell1,
            "<p><span style='color:red'>Red</span> " +
            "<span style='color:#0000FF'>Blue</span> " +
            "<span style='color:rgb(0,255,0)'>Green</span></p>");

        // Font family examples
        XSSFCell cell2 = sheet.createRow(rowNum++).createCell(0);
        converter.applyHtmlToCell(cell2,
            "<p><span style='font-family:Arial'>Arial</span> " +
            "<span style='font-family:Courier New'>Courier</span></p>");

        // Font size examples
        XSSFCell cell3 = sheet.createRow(rowNum++).createCell(0);
        converter.applyHtmlToCell(cell3,
            "<p><span style='font-size:10px'>Small</span> " +
            "<span style='font-size:14px'>Medium</span> " +
            "<span style='font-size:18px'>Large</span></p>");

        // Combined styles
        XSSFCell cell4 = sheet.createRow(rowNum++).createCell(0);
        converter.applyHtmlToCell(cell4,
            "<p><b><i><u><span style='color:red;font-size:16px'>Bold Italic Underline Red 16px</span></u></i></b></p>");

        sheet.setColumnWidth(0, 12000);

        try (FileOutputStream fos = new FileOutputStream("example2-richtext.xlsx")) {
            workbook.write(fos);
        }
        workbook.close();

        System.out.println("✓ Example 2: Rich text styles -> example2-richtext.xlsx");
    }

    /**
     * Example 3: Lists (ordered and unordered)
     */
    public static void listExample() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Lists");
        HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);

        // Unordered list
        XSSFCell cell1 = sheet.createRow(0).createCell(0);
        converter.applyHtmlToCell(cell1,
            "<ul>" +
            "  <li>First item</li>" +
            "  <li>Second item</li>" +
            "  <li>Third item</li>" +
            "</ul>");
        sheet.createRow(0).createCell(1).setCellValue("Unordered List");

        // Ordered list
        XSSFCell cell2 = sheet.createRow(5).createCell(0);
        converter.applyHtmlToCell(cell2,
            "<ol>" +
            "  <li>Step one</li>" +
            "  <li>Step two</li>" +
            "  <li>Step three</li>" +
            "</ol>");
        sheet.createRow(5).createCell(1).setCellValue("Ordered List");

        // Styled list
        XSSFCell cell3 = sheet.createRow(10).createCell(0);
        converter.applyHtmlToCell(cell3,
            "<ul>" +
            "  <li><b>Bold item</b></li>" +
            "  <li><span style='color:red'>Red item</span></li>" +
            "  <li><i>Italic item</i></li>" +
            "</ul>");
        sheet.createRow(10).createCell(1).setCellValue("Styled List");

        sheet.setColumnWidth(0, 8000);
        sheet.setColumnWidth(1, 5000);

        try (FileOutputStream fos = new FileOutputStream("example3-lists.xlsx")) {
            workbook.write(fos);
        }
        workbook.close();

        System.out.println("✓ Example 3: Lists -> example3-lists.xlsx");
    }

    /**
     * Example 4: Tables
     */
    public static void tableExample() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Tables");
        HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);

        // Simple table
        XSSFCell cell1 = sheet.createRow(0).createCell(0);
        converter.applyHtmlToCell(cell1,
            "<table>" +
            "  <tr><td>Name</td><td>Age</td><td>City</td></tr>" +
            "  <tr><td>John</td><td>25</td><td>New York</td></tr>" +
            "  <tr><td>Jane</td><td>30</td><td>London</td></tr>" +
            "</table>");
        sheet.createRow(0).createCell(1).setCellValue("Simple Table");

        // Styled table
        XSSFCell cell2 = sheet.createRow(5).createCell(0);
        converter.applyHtmlToCell(cell2,
            "<table>" +
            "  <tr><td><b>Product</b></td><td><b>Price</b></td></tr>" +
            "  <tr><td>Apple</td><td><span style='color:green'>$1.50</span></td></tr>" +
            "  <tr><td>Banana</td><td><span style='color:green'>$0.80</span></td></tr>" +
            "</table>");
        sheet.createRow(5).createCell(1).setCellValue("Styled Table");

        sheet.setColumnWidth(0, 12000);
        sheet.setColumnWidth(1, 5000);

        try (FileOutputStream fos = new FileOutputStream("example4-tables.xlsx")) {
            workbook.write(fos);
        }
        workbook.close();

        System.out.println("✓ Example 4: Tables -> example4-tables.xlsx");
    }

    /**
     * Example 5: Hyperlinks
     */
    public static void hyperlinkExample() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Hyperlinks");
        HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);

        XSSFCell cell = sheet.createRow(0).createCell(0);
        converter.applyHtmlToCell(cell,
            "<p>" +
            "<a href='https://github.com'>GitHub</a><br/>" +
            "<a href='https://stackoverflow.com'>Stack Overflow</a><br/>" +
            "<a href='mailto:example@example.com'>Email Us</a>" +
            "</p>");

        sheet.setColumnWidth(0, 8000);

        try (FileOutputStream fos = new FileOutputStream("example5-hyperlinks.xlsx")) {
            workbook.write(fos);
        }
        workbook.close();

        System.out.println("✓ Example 5: Hyperlinks -> example5-hyperlinks.xlsx");
    }

    /**
     * Example 6: Background colors
     */
    public static void backgroundExample() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Backgrounds");
        HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);

        int rowNum = 0;

        // Yellow background
        XSSFCell cell1 = sheet.createRow(rowNum++).createCell(0);
        converter.applyHtmlToCell(cell1,
            "<p style='background-color:#FFFF00'>Yellow highlight</p>");

        // Light blue background
        XSSFCell cell2 = sheet.createRow(rowNum++).createCell(0);
        converter.applyHtmlToCell(cell2,
            "<p style='background-color:#ADD8E6'>Light blue highlight</p>");

        // Light green background with colored text
        XSSFCell cell3 = sheet.createRow(rowNum++).createCell(0);
        converter.applyHtmlToCell(cell3,
            "<p style='background-color:#90EE90'><b><span style='color:darkgreen'>Green theme</span></b></p>");

        sheet.setColumnWidth(0, 8000);

        try (FileOutputStream fos = new FileOutputStream("example6-backgrounds.xlsx")) {
            workbook.write(fos);
        }
        workbook.close();

        System.out.println("✓ Example 6: Backgrounds -> example6-backgrounds.xlsx");
    }

    /**
     * Example 7: Images with async download
     */
    public static void imageExample() throws Exception {
        ConverterConfig config = ConverterConfig.builder()
            .enableImageDownload(true)
            .imageTimeout(5000, 15000)
            .build();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Images");
        HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook, config);

        XSSFCell cell = sheet.createRow(0).createCell(0);

        // Note: This will attempt to download the image
        // Use a valid image URL or comment out if offline
        String html = "<p><b>Image example</b></p>" +
                     "<img src='https://via.placeholder.com/150'/>";

        converter.applyHtmlToCell(cell, html);

        sheet.setColumnWidth(0, 8000);
        sheet.createRow(0).setHeightInPoints(120);

        try (FileOutputStream fos = new FileOutputStream("example7-images.xlsx")) {
            workbook.write(fos);
        }
        workbook.close();

        System.out.println("✓ Example 7: Images (if online) -> example7-images.xlsx");
    }
}
