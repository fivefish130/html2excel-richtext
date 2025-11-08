package io.github.fivefish130.html2excel.richtext.cache;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Cell style cache to avoid creating duplicate style objects
 * Excel has limits on cell styles (~64K), caching helps control this
 *
 * @author fivefish130
 */
public class StyleCache {

    private final XSSFWorkbook workbook;
    private final Map<String, XSSFCellStyle> cache;
    private final boolean enabled;

    public StyleCache(XSSFWorkbook workbook, boolean enabled) {
        this.workbook = workbook;
        this.enabled = enabled;
        this.cache = enabled ? new ConcurrentHashMap<>() : null;
    }

    /**
     * Get or create cell style with given key
     *
     * @param key Style cache key
     * @param creator Style creator function
     * @return Cached or newly created style
     */
    public XSSFCellStyle getOrCreate(String key, StyleCreator creator) {
        if (!enabled) {
            return creator.create(workbook);
        }

        return cache.computeIfAbsent(key, k -> creator.create(workbook));
    }

    /**
     * Generate style cache key for background color
     */
    public static String generateBackgroundKey(String colorStr) {
        return "bg:" + (colorStr != null ? colorStr.trim().toLowerCase() : "none");
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
     * Functional interface for style creation
     */
    @FunctionalInterface
    public interface StyleCreator {
        XSSFCellStyle create(XSSFWorkbook workbook);
    }
}
