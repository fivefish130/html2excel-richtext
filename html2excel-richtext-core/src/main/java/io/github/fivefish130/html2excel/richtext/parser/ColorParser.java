package io.github.fivefish130.html2excel.richtext.parser;

import java.awt.Color;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Color parser supporting multiple formats
 * - Hex: #FF0000, #F00
 * - RGB: rgb(255,0,0)
 * - Named: red, blue, green, etc.
 *
 * @author fivefish130
 */
public class ColorParser {

    private static final Pattern RGB_PATTERN = Pattern.compile(
            "rgb\\s*\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)\\s*\\)");

    /**
     * Parse color string to Color object
     *
     * @param input Color string (#hex, rgb(), or named color)
     * @return Color object, or null if parsing fails
     */
    public static Color parse(String input) {
        if (input == null) {
            return null;
        }

        String s = input.trim().toLowerCase(Locale.ROOT);
        try {
            // Hex format: #RGB or #RRGGBB
            if (s.startsWith("#")) {
                return Color.decode(s);
            }

            // RGB function: rgb(r, g, b)
            if (s.startsWith("rgb")) {
                Matcher matcher = RGB_PATTERN.matcher(s);
                if (matcher.find()) {
                    int r = Integer.parseInt(matcher.group(1));
                    int g = Integer.parseInt(matcher.group(2));
                    int b = Integer.parseInt(matcher.group(3));
                    return new Color(r, g, b);
                }
            }

            // Named colors
            return parseNamedColor(s);

        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Parse named color to Color object
     */
    private static Color parseNamedColor(String name) {
        switch (name) {
            case "black": return Color.BLACK;
            case "white": return Color.WHITE;
            case "red": return Color.RED;
            case "green": return Color.GREEN;
            case "blue": return Color.BLUE;
            case "yellow": return Color.YELLOW;
            case "gray":
            case "grey": return Color.GRAY;
            case "orange": return Color.ORANGE;
            case "purple": return new Color(128, 0, 128);
            case "brown": return new Color(165, 42, 42);
            case "pink": return Color.PINK;
            case "cyan": return Color.CYAN;
            case "magenta": return Color.MAGENTA;
            default: return null;
        }
    }
}
