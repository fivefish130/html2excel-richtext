package io.github.fivefish130.html2excel.richtext.parser;

import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * Pure Java CSS declaration parser with high fault tolerance
 *
 * Supported formats:
 * - "color: red; font-size: 14px; background-color: #FFFF00"
 * - "color:red;font-size:14px" (no spaces)
 * - "color: rgb(255,0,0); font-weight: bold"
 *
 * @author fivefish130
 */
public class CssParser {

    /**
     * Parse CSS declaration string
     *
     * @param cssText CSS declaration (e.g., "color:red; font-size:14px")
     * @return Map of CSS properties
     */
    public Map<String, String> parseDeclaration(String cssText) {
        Map<String, String> result = new HashMap<>();

        if (cssText == null || cssText.trim().isEmpty()) {
            return result;
        }

        try {
            // Split by semicolon
            String[] declarations = cssText.split(";");
            for (String declaration : declarations) {
                declaration = declaration.trim();
                if (declaration.isEmpty()) {
                    continue;
                }

                // Find first colon
                int colonIndex = declaration.indexOf(':');
                if (colonIndex == -1) {
                    continue;
                }

                String propName = declaration.substring(0, colonIndex)
                        .trim()
                        .toLowerCase(Locale.ROOT);
                String propValue = declaration.substring(colonIndex + 1).trim();

                // Remove quotes
                propValue = propValue.replaceAll("^['\"]|['\"]$", "");

                if (!propName.isEmpty() && !propValue.isEmpty()) {
                    result.put(propName, propValue);
                }
            }
        } catch (Exception e) {
            // Fault tolerance: ignore parsing errors
        }

        return result;
    }

    /**
     * Normalize font name (take first font, remove quotes)
     */
    public String normalizeFontName(String name) {
        if (name == null) {
            return null;
        }
        String[] parts = name.split(",");
        return parts[0].replaceAll("[\"']", "").trim();
    }

    /**
     * Parse font size to points
     * Supports: "12px", "12pt", "14", "14.0pt"
     */
    public Short parseFontSize(String raw) {
        if (raw == null || raw.isEmpty()) {
            return null;
        }

        try {
            // Extract digits
            String digits = raw.replaceAll("[^0-9.]", "");
            if (digits.isEmpty()) {
                return null;
            }

            float val = Float.parseFloat(digits);

            // Convert px to pt (1px ≈ 0.75pt)
            if (raw.endsWith("px")) {
                val = val * 0.75f;
            }

            return (short) Math.max(8, Math.round(val));
        } catch (Exception e) {
            return null;
        }
    }
}
