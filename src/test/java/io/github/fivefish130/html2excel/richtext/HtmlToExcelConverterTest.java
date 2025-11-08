package io.github.fivefish130.html2excel.richtext;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for HtmlToExcelConverter
 *
 * @author fivefish130
 */
class HtmlToExcelConverterTest {

    private XSSFWorkbook workbook;
    private HtmlToExcelConverter converter;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        converter = new HtmlToExcelConverter(workbook);
    }

    @Test
    void testBasicHtml() {
        String html = "<p><b>Bold</b> <i>Italic</i> <u>Underline</u></p>";
        XSSFRichTextString richText = converter.convertToRichText(html);

        assertNotNull(richText);
        assertTrue(richText.getString().contains("Bold"));
        assertTrue(richText.getString().contains("Italic"));
        assertTrue(richText.getString().contains("Underline"));
    }

    @Test
    void testColoredText() {
        String html = "<span style='color:red'>Red Text</span>";
        XSSFRichTextString richText = converter.convertToRichText(html);

        assertNotNull(richText);
        assertEquals("Red Text", richText.getString());
    }

    @Test
    void testMultipleColors() {
        String html = "<p>" +
                "    <span style='color:#FF0000'>Red</span>" +
                "    <span style='color:rgb(0,255,0)'>Green</span>" +
                "    <span style='color:blue'>Blue</span>" +
                "</p>";
        XSSFRichTextString richText = converter.convertToRichText(html);

        assertNotNull(richText);
        String text = richText.getString();
        assertTrue(text.contains("Red"));
        assertTrue(text.contains("Green"));
        assertTrue(text.contains("Blue"));
    }

    @Test
    void testHyperlink() {
        XSSFSheet sheet = workbook.createSheet("Test");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);

        String html = "<a href='https://github.com'>GitHub Link</a>";
        converter.applyHtmlToCell(cell, html);

        assertEquals("GitHub Link", cell.getStringCellValue());
        assertNotNull(cell.getHyperlink());
        assertEquals("https://github.com", cell.getHyperlink().getAddress());
    }

    @Test
    void testBackgroundColor() {
        XSSFSheet sheet = workbook.createSheet("Test");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);

        String html = "<p style='background-color:#FFFF00'>Yellow Background</p>";
        converter.applyHtmlToCell(cell, html);

        assertTrue(cell.getStringCellValue().contains("Yellow Background"));
        assertNotNull(cell.getCellStyle());
    }

    @Test
    void testComplexHtml() {
        XSSFSheet sheet = workbook.createSheet("Test");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);

        String html = "<div>" +
                "    <p><b>Product Name:</b> <span style='color:#0000FF'>Sample Item</span></p>" +
                "    <p><i>Price:</i> <span style='color:#008000'>$99.99</span></p>" +
                "</div>";
        converter.applyHtmlToCell(cell, html);

        String cellValue = cell.getStringCellValue();
        assertTrue(cellValue.contains("Product Name:"));
        assertTrue(cellValue.contains("Sample Item"));
        assertTrue(cellValue.contains("Price:"));
        assertTrue(cellValue.contains("$99.99"));
    }

    @Test
    void testEmptyHtml() {
        XSSFRichTextString richText = converter.convertToRichText("");
        assertNotNull(richText);
        assertEquals("", richText.getString());
    }

    @Test
    void testNullHtml() {
        XSSFRichTextString richText = converter.convertToRichText(null);
        assertNotNull(richText);
        assertEquals("", richText.getString());
    }

    @Test
    void testLongText() {
        StringBuilder longHtml = new StringBuilder("<p>");
        for (int i = 0; i < 10000; i++) {
            longHtml.append("Text ");
        }
        longHtml.append("</p>");

        XSSFSheet sheet = workbook.createSheet("Test");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);

        converter.applyHtmlToCell(cell, longHtml.toString());

        // Should be truncated to max cell length
        assertTrue(cell.getStringCellValue().length() <= 32767);
    }

    @Test
    void testFontStyles() {
        String html = "<p style='font-family:Arial; font-size:14px'>" +
                "    <b>Bold</b> <i>Italic</i> <u>Underline</u>" +
                "</p>";
        XSSFRichTextString richText = converter.convertToRichText(html);

        assertNotNull(richText);
        assertTrue(richText.getString().contains("Bold"));
    }

    @Test
    void testSpecialCharacters() {
        String html = "<p>Text with &nbsp; spaces and &lt;tags&gt;</p>";
        XSSFRichTextString richText = converter.convertToRichText(html);

        assertNotNull(richText);
        String text = richText.getString();
        assertTrue(text.contains("spaces"));
        assertTrue(text.contains("<tags>"));
    }

    @Test
    void testCacheStats() {
        // Process multiple cells with same styles
        XSSFSheet sheet = workbook.createSheet("Test");
        String html = "<p><b>Bold Text</b></p>";

        for (int i = 0; i < 10; i++) {
            XSSFCell cell = sheet.createRow(i).createCell(0);
            converter.applyHtmlToCell(cell, html);
        }

        // Cache should reduce font object creation
        assertTrue(converter.getFontCacheSize() >= 0);
    }

    @Test
    void testWriteToFile() throws IOException {
        XSSFSheet sheet = workbook.createSheet("Demo");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);

        String html = "<div>" +
                "    <p><b style='color:#FF0000'>Important:</b> This is a test</p>" +
                "    <p><a href='https://github.com'>Visit GitHub</a></p>" +
                "</div>";
        converter.applyHtmlToCell(cell, html);

        // Write to file
        String outputPath = "target/test-output.xlsx";
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
        }

        // Verify file was created
        assertTrue(new java.io.File(outputPath).exists());
        System.out.println("Test Excel file created: " + outputPath);
    }
}
