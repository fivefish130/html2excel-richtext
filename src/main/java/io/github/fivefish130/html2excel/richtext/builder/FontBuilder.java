package io.github.fivefish130.html2excel.richtext.builder;

import io.github.fivefish130.html2excel.richtext.cache.FontCache;
import io.github.fivefish130.html2excel.richtext.parser.ColorParser;
import io.github.fivefish130.html2excel.richtext.parser.CssParser;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;
import java.util.Map;

/**
 * Font builder to create XSSFFont from style properties
 *
 * @author fivefish130
 */
public class FontBuilder {

    private static final Logger log = LoggerFactory.getLogger(FontBuilder.class);

    private final XSSFWorkbook workbook;
    private final FontCache fontCache;
    private final CssParser cssParser;

    public FontBuilder(XSSFWorkbook workbook, FontCache fontCache) {
        this.workbook = workbook;
        this.fontCache = fontCache;
        this.cssParser = new CssParser();
    }

    /**
     * Build or get cached font from style properties
     */
    public XSSFFont buildFont(Map<String, String> style) {
        String cacheKey = FontCache.generateKey(style);
        return fontCache.getOrCreate(cacheKey, wb -> createFont(wb, style));
    }

    /**
     * Create new font from style properties
     */
    private XSSFFont createFont(XSSFWorkbook wb, Map<String, String> style) {
        XSSFFont font = wb.createFont();

        // Font family
        if (style.containsKey("font-family")) {
            String fontName = cssParser.normalizeFontName(style.get("font-family"));
            if (fontName != null && !fontName.isEmpty()) {
                font.setFontName(fontName);
            }
        }

        // Font size
        if (style.containsKey("font-size")) {
            Short fontSize = cssParser.parseFontSize(style.get("font-size"));
            if (fontSize != null) {
                font.setFontHeightInPoints(fontSize);
            }
        }

        // Bold
        if ("bold".equalsIgnoreCase(style.get("font-weight"))) {
            font.setBold(true);
        }

        // Italic
        if ("italic".equalsIgnoreCase(style.get("font-style"))) {
            font.setItalic(true);
        }

        // Underline
        if ("underline".equalsIgnoreCase(style.get("text-decoration"))) {
            font.setUnderline(FontUnderline.SINGLE);
        }

        // Color
        if (style.containsKey("color")) {
            Color c = ColorParser.parse(style.get("color"));
            if (c != null) {
                try {
                    XSSFColor xssfColor = new XSSFColor(c, wb.getStylesSource().getIndexedColors());
                    font.setColor(xssfColor);
                } catch (Exception e) {
                    log.warn("Failed to set font color: {}", e.getMessage());
                }
            }
        }

        return font;
    }
}
