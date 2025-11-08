package io.github.fivefish130.html2excel.richtext.handler;

import io.github.fivefish130.html2excel.richtext.cache.StyleCache;
import io.github.fivefish130.html2excel.richtext.parser.ColorParser;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;

/**
 * Handler for cell background color
 *
 * @author fivefish130
 */
public class BackgroundHandler {

    private final XSSFWorkbook workbook;
    private final StyleCache styleCache;

    public BackgroundHandler(XSSFWorkbook workbook, StyleCache styleCache) {
        this.workbook = workbook;
        this.styleCache = styleCache;
    }

    /**
     * Apply background color to cell (only if not already set)
     *
     * @param cell Target cell
     * @param colorStr Background color string
     */
    public void applyBackground(XSSFCell cell, String colorStr) {
        if (colorStr == null || colorStr.trim().isEmpty()) {
            return;
        }

        // Skip if cell already has background color
        XSSFCellStyle existing = cell.getCellStyle();
        if (existing != null && existing.getFillPattern() != null &&
            existing.getFillForegroundColorColor() != null) {
            return;
        }

        // Get or create cached style
        String cacheKey = StyleCache.generateBackgroundKey(colorStr);
        XSSFCellStyle style = styleCache.getOrCreate(cacheKey, wb -> createBackgroundStyle(wb, colorStr));

        cell.setCellStyle(style);
    }

    /**
     * Create cell style with background color
     */
    private XSSFCellStyle createBackgroundStyle(XSSFWorkbook wb, String colorStr) {
        XSSFCellStyle style = wb.createCellStyle();

        Color color = ColorParser.parse(colorStr);
        if (color != null) {
            XSSFColor xssfColor = new XSSFColor(color, wb.getStylesSource().getIndexedColors());
            style.setFillForegroundColor(xssfColor);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        style.setVerticalAlignment(VerticalAlignment.TOP);

        return style;
    }
}
