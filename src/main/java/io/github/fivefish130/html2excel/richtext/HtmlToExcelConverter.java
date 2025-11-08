package io.github.fivefish130.html2excel.richtext;

import io.github.fivefish130.html2excel.richtext.builder.FontBuilder;
import io.github.fivefish130.html2excel.richtext.cache.FontCache;
import io.github.fivefish130.html2excel.richtext.cache.StyleCache;
import io.github.fivefish130.html2excel.richtext.config.ConverterConfig;
import io.github.fivefish130.html2excel.richtext.handler.BackgroundHandler;
import io.github.fivefish130.html2excel.richtext.handler.HyperlinkHandler;
import io.github.fivefish130.html2excel.richtext.handler.ImageHandler;
import io.github.fivefish130.html2excel.richtext.parser.HtmlTraverser;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.util.HashMap;
import java.util.Objects;

/**
 * HTML to Excel Rich Text Converter (Facade)
 * <p>
 * Features:
 * - Complete rich text support (bold, italic, underline, colors, fonts, sizes)
 * - Pure Java CSS parser (no external dependencies)
 * - Cell background color mapping
 * - Hyperlink auto-extraction
 * - Image download and embedding
 * - High fault tolerance (Jsoup auto-fixes malformed HTML)
 * - Font/Style caching for performance
 * - Long text auto-truncation
 * <p>
 * Example:
 * <pre>
 * XSSFWorkbook workbook = new XSSFWorkbook();
 * XSSFCell cell = ...;
 *
 * HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook);
 * converter.applyHtmlToCell(cell, "&lt;p&gt;&lt;b&gt;Bold&lt;/b&gt; &lt;i&gt;Italic&lt;/i&gt;&lt;/p&gt;");
 * </pre>
 *
 * @author fivefish130
 * @since 1.0.0
 */
public class HtmlToExcelConverter {

    private final XSSFWorkbook workbook;
    private final ConverterConfig config;

    // Core components
    private final FontCache fontCache;
    private final StyleCache styleCache;
    private final FontBuilder fontBuilder;
    private final BackgroundHandler backgroundHandler;
    private final HyperlinkHandler hyperlinkHandler;
    private final ImageHandler imageHandler;
    private final HtmlTraverser htmlTraverser;

    /**
     * Create converter with default configuration
     *
     * @param workbook Excel workbook
     */
    public HtmlToExcelConverter(XSSFWorkbook workbook) {
        this(workbook, new ConverterConfig());
    }

    /**
     * Create converter with custom configuration
     *
     * @param workbook Excel workbook
     * @param config Converter configuration
     */
    public HtmlToExcelConverter(XSSFWorkbook workbook, ConverterConfig config) {
        this.workbook = Objects.requireNonNull(workbook, "workbook cannot be null");
        this.config = Objects.requireNonNull(config, "config cannot be null");

        // Initialize components
        this.fontCache = new FontCache(workbook, config.isEnableFontCache());
        this.styleCache = new StyleCache(workbook, config.isEnableStyleCache());
        this.fontBuilder = new FontBuilder(workbook, fontCache);
        this.backgroundHandler = new BackgroundHandler(workbook, styleCache);
        this.hyperlinkHandler = new HyperlinkHandler(workbook);
        this.imageHandler = new ImageHandler(config);
        this.htmlTraverser = new HtmlTraverser(fontBuilder, backgroundHandler);
    }

    /**
     * Convert HTML to rich text string (text only, no cell styling)
     *
     * @param html HTML string
     * @return Rich text string with formatting
     */
    public XSSFRichTextString convertToRichText(String html) {
        Document doc = Jsoup.parseBodyFragment(defaultIfNull(html));
        Element body = doc.body();

        XSSFRichTextString rich = new XSSFRichTextString("");
        htmlTraverser.traverse(body, new HashMap<>(), rich, null);
        return rich;
    }

    /**
     * Apply HTML to cell (includes rich text, hyperlink, background, images)
     *
     * @param cell Target cell
     * @param html HTML string
     */
    public void applyHtmlToCell(XSSFCell cell, String html) {
        if (cell == null) {
            throw new IllegalArgumentException("cell cannot be null");
        }

        Document doc = Jsoup.parseBodyFragment(defaultIfNull(html));
        Element body = doc.body();

        // 1. Find first hyperlink
        String firstHref = hyperlinkHandler.findFirstHref(body);

        // 2. Build rich text
        XSSFRichTextString rich = new XSSFRichTextString("");
        htmlTraverser.traverse(body, new HashMap<>(), rich, cell);

        // 3. Set cell value (handle long text)
        String fullText = rich.getString();
        if (fullText.length() > config.getMaxCellLength()) {
            int maxLength = config.getMaxCellLength() - config.getTruncateSuffix().length();
            String truncated = fullText.substring(0, maxLength) + config.getTruncateSuffix();
            cell.setCellValue(truncated);
        } else {
            cell.setCellValue(rich);
        }

        // 4. Apply hyperlink
        if (firstHref != null && !firstHref.trim().isEmpty()) {
            hyperlinkHandler.applyHyperlink(cell, firstHref);
        }

        // 5. Process images
        if (config.isEnableImageDownload()) {
            imageHandler.processImages(body, cell);
        }
    }

    /**
     * Set image download timeout
     *
     * @param connectTimeout Connect timeout in milliseconds
     * @param readTimeout Read timeout in milliseconds
     * @deprecated Use {@link ConverterConfig.Builder#imageTimeout(int, int)} instead
     */
    @Deprecated
    public void setImageDownloadTimeout(int connectTimeout, int readTimeout) {
        config.setImageConnectTimeout(connectTimeout);
        config.setImageReadTimeout(readTimeout);
    }

    /**
     * Enable or disable image download
     *
     * @param enable true to enable image download
     * @deprecated Use {@link ConverterConfig.Builder#enableImageDownload(boolean)} instead
     */
    @Deprecated
    public void setEnableImageDownload(boolean enable) {
        config.setEnableImageDownload(enable);
    }

    /**
     * Get current configuration
     *
     * @return Converter configuration
     */
    public ConverterConfig getConfig() {
        return config;
    }

    /**
     * Get font cache statistics
     *
     * @return Number of cached fonts
     */
    public int getFontCacheSize() {
        return fontCache.size();
    }

    /**
     * Get style cache statistics
     *
     * @return Number of cached styles
     */
    public int getStyleCacheSize() {
        return styleCache.size();
    }

    /**
     * Clear all caches
     */
    public void clearCaches() {
        fontCache.clear();
        styleCache.clear();
    }

    private static String defaultIfNull(String s) {
        return s == null ? "" : s;
    }
}
