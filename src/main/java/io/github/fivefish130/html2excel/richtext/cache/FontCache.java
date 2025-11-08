package io.github.fivefish130.html2excel.richtext.cache;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Font cache to avoid creating duplicate font objects
 * Excel has limits on the number of fonts (~64K), caching helps control this
 *
 * @author fivefish130
 */
public class FontCache {

    private final XSSFWorkbook workbook;
    private final Map<String, XSSFFont> cache;
    private final boolean enabled;

    public FontCache(XSSFWorkbook workbook, boolean enabled) {
        this.workbook = workbook;
        this.enabled = enabled;
        this.cache = enabled ? new ConcurrentHashMap<>() : null;
    }

    /**
     * Get or create font with given key
     *
     * @param key Font cache key
     * @param creator Font creator function
     * @return Cached or newly created font
     */
    public XSSFFont getOrCreate(String key, FontCreator creator) {
        if (!enabled) {
            return creator.create(workbook);
        }

        return cache.computeIfAbsent(key, k -> creator.create(workbook));
    }

    /**
     * Generate font cache key from style properties
     */
    public static String generateKey(Map<String, String> style) {
        StringBuilder key = new StringBuilder();

        // Font family
        String fontFamily = style.get("font-family");
        key.append("family:").append(fontFamily != null ? fontFamily : "default").append("|");

        // Font size
        String fontSize = style.get("font-size");
        key.append("size:").append(fontSize != null ? fontSize : "default").append("|");

        // Font weight
        String fontWeight = style.get("font-weight");
        key.append("weight:").append("bold".equalsIgnoreCase(fontWeight) ? "bold" : "normal").append("|");

        // Font style
        String fontStyle = style.get("font-style");
        key.append("style:").append("italic".equalsIgnoreCase(fontStyle) ? "italic" : "normal").append("|");

        // Text decoration
        String textDecoration = style.get("text-decoration");
        key.append("decoration:").append("underline".equalsIgnoreCase(textDecoration) ? "underline" : "none").append("|");

        // Color
        String color = style.get("color");
        key.append("color:").append(color != null ? color : "default");

        return key.toString();
    }

    /**
     * Clear cache
     */
    public void clear() {
        if (cache != null) {
            cache.clear();
        }
    }

    /**
     * Get cache size
     */
    public int size() {
        return cache != null ? cache.size() : 0;
    }

    /**
     * Functional interface for font creation
     */
    @FunctionalInterface
    public interface FontCreator {
        XSSFFont create(XSSFWorkbook workbook);
    }
}
