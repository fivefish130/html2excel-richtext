package io.github.fivefish130.html2excel.richtext.parser;

import io.github.fivefish130.html2excel.richtext.builder.FontBuilder;
import io.github.fivefish130.html2excel.richtext.handler.BackgroundHandler;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;

import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * HTML DOM traverser to build rich text string
 *
 * @author fivefish130
 */
public class HtmlTraverser {

    private final FontBuilder fontBuilder;
    private final BackgroundHandler backgroundHandler;
    private final CssParser cssParser;

    public HtmlTraverser(FontBuilder fontBuilder, BackgroundHandler backgroundHandler) {
        this.fontBuilder = fontBuilder;
        this.backgroundHandler = backgroundHandler;
        this.cssParser = new CssParser();
    }

    /**
     * Traverse HTML node and build rich text
     *
     * @param node Current node
     * @param inheritedStyle Inherited CSS styles
     * @param rich Rich text string being built
     * @param targetCell Target cell (for background color)
     */
    public void traverse(Node node, Map<String, String> inheritedStyle,
                        XSSFRichTextString rich, XSSFCell targetCell) {
        traverse(node, inheritedStyle, rich, targetCell, new TraverseContext());
    }

    private void traverse(Node node, Map<String, String> inheritedStyle,
                         XSSFRichTextString rich, XSSFCell targetCell, TraverseContext context) {
        // Handle text nodes
        if (node instanceof TextNode) {
            TextNode textNode = (TextNode) node;
            String txt = textNode.getWholeText()
                    .replace('\u00A0', ' ')  // &nbsp;
                    .replaceAll("[\\t\\r\\f\\v]+", " ");
            if (!txt.isEmpty()) {
                rich.append(txt);
            }
            return;
        }

        // Handle element nodes
        if (!(node instanceof Element)) {
            return;
        }

        Element el = (Element) node;
        String tag = el.tagName().toLowerCase(Locale.ROOT);

        // Handle <br> explicitly
        if ("br".equals(tag)) {
            rich.append("\n");
            return;
        }

        // Merge styles
        Map<String, String> style = new HashMap<>(inheritedStyle);
        updateStyleFromTag(tag, style);
        updateStyleFromAttr(el, style);

        // Handle list items with bullets/numbers
        if ("li".equals(tag)) {
            handleListItem(el, style, rich, targetCell, context);
            return;
        }

        // Handle table rows
        if ("tr".equals(tag)) {
            handleTableRow(el, style, rich, targetCell, context);
            return;
        }

        // Handle table cells
        if ("td".equals(tag) || "th".equals(tag)) {
            handleTableCell(el, style, rich, targetCell, context);
            return;
        }

        // Handle block elements
        if (isBlockTag(tag)) {
            int blockStart = rich.length();

            // Track list context
            if ("ul".equals(tag)) {
                context.enterList(false);
            } else if ("ol".equals(tag)) {
                context.enterList(true);
            }

            // Traverse children
            for (Node child : el.childNodes()) {
                traverse(child, style, rich, targetCell, context);
            }

            // Exit list context
            if ("ul".equals(tag) || "ol".equals(tag)) {
                context.exitList();
            }

            int blockEnd = rich.length();

            // Add newline only if block has content
            if (blockEnd > blockStart) {
                rich.append("\n");
            }

            // Apply background color to cell
            if (targetCell != null && style.containsKey("background-color")) {
                backgroundHandler.applyBackground(targetCell, style.get("background-color"));
            }
            return;
        }

        // Handle inline elements
        int start = rich.length();
        for (Node child : el.childNodes()) {
            traverse(child, style, rich, targetCell, context);
        }
        int end = rich.length();

        // Apply font only if there's content and effective style changes
        if (end > start && hasEffectiveStyle(style, inheritedStyle)) {
            XSSFFont font = fontBuilder.buildFont(style);
            if (font != null) {
                rich.applyFont(start, end, font);
            }
        }
    }

    /**
     * Handle list item with bullet or number
     */
    private void handleListItem(Element el, Map<String, String> style,
                                XSSFRichTextString rich, XSSFCell targetCell, TraverseContext context) {
        // Add bullet or number
        if (context.isOrderedList()) {
            int itemNumber = context.getAndIncrementItemNumber();
            rich.append(itemNumber + ". ");
        } else {
            rich.append("\u2022 ");  // Bullet point: •
        }

        // Traverse content
        for (Node child : el.childNodes()) {
            traverse(child, style, rich, targetCell, context);
        }

        rich.append("\n");
    }

    /**
     * Handle table row
     */
    private void handleTableRow(Element el, Map<String, String> style,
                                XSSFRichTextString rich, XSSFCell targetCell, TraverseContext context) {
        context.enterRow();

        for (Node child : el.childNodes()) {
            traverse(child, style, rich, targetCell, context);
        }

        context.exitRow();
        rich.append("\n");
    }

    /**
     * Handle table cell
     */
    private void handleTableCell(Element el, Map<String, String> style,
                                 XSSFRichTextString rich, XSSFCell targetCell, TraverseContext context) {
        // Add separator for non-first cells
        if (context.getCellIndex() > 0) {
            rich.append(" | ");
        }

        context.incrementCellIndex();

        // Traverse content
        for (Node child : el.childNodes()) {
            traverse(child, style, rich, targetCell, context);
        }
    }

    /**
     * Context for traversing (tracks list/table state)
     */
    private static class TraverseContext {
        private boolean inOrderedList = false;
        private int listItemNumber = 1;
        private int cellIndex = 0;

        void enterList(boolean ordered) {
            this.inOrderedList = ordered;
            this.listItemNumber = 1;
        }

        void exitList() {
            this.inOrderedList = false;
            this.listItemNumber = 1;
        }

        boolean isOrderedList() {
            return inOrderedList;
        }

        int getAndIncrementItemNumber() {
            return listItemNumber++;
        }

        void enterRow() {
            this.cellIndex = 0;
        }

        void exitRow() {
            this.cellIndex = 0;
        }

        int getCellIndex() {
            return cellIndex;
        }

        void incrementCellIndex() {
            cellIndex++;
        }
    }

    /**
     * Check if current style has effective changes compared to inherited
     */
    private boolean hasEffectiveStyle(Map<String, String> currentStyle,
                                     Map<String, String> inheritedStyle) {
        String[] keyProperties = {
                "font-weight", "font-style", "text-decoration",
                "color", "font-family", "font-size"
        };

        for (String prop : keyProperties) {
            String currentValue = currentStyle.get(prop);
            String inheritedValue = inheritedStyle.get(prop);

            if (currentValue != null && !currentValue.equals(inheritedValue)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Check if tag is block-level element
     */
    private boolean isBlockTag(String tag) {
        switch (tag) {
            case "p":
            case "div":
            case "h1":
            case "h2":
            case "h3":
            case "h4":
            case "h5":
            case "h6":
            case "li":
            case "ul":
            case "ol":
            case "table":
            case "tr":
            case "blockquote":
                return true;
            default:
                return false;
        }
    }

    /**
     * Update style map from HTML tag
     */
    private void updateStyleFromTag(String tag, Map<String, String> style) {
        switch (tag) {
            case "b":
            case "strong":
                style.put("font-weight", "bold");
                break;
            case "i":
            case "em":
                style.put("font-style", "italic");
                break;
            case "u":
                style.put("text-decoration", "underline");
                break;
            case "a":
                style.put("color", "#0563C1");  // Excel link color
                style.put("text-decoration", "underline");
                break;
            case "code":
                style.put("font-family", "Courier New");
                break;
            default:
                break;
        }
    }

    /**
     * Update style map from HTML attributes
     */
    private void updateStyleFromAttr(Element el, Map<String, String> style) {
        // HTML attributes
        if (el.hasAttr("color")) {
            style.put("color", el.attr("color"));
        }
        if (el.hasAttr("bgcolor")) {
            style.put("background-color", el.attr("bgcolor"));
        }

        // Parse style attribute
        if (el.hasAttr("style")) {
            String cssText = el.attr("style");
            Map<String, String> parsed = cssParser.parseDeclaration(cssText);
            style.putAll(parsed);
        }

        // Handle <font> tag
        if (el.tagName().equalsIgnoreCase("font")) {
            if (el.hasAttr("face")) {
                style.put("font-family", el.attr("face"));
            }
            if (el.hasAttr("size")) {
                try {
                    int v = Integer.parseInt(el.attr("size"));
                    style.put("font-size", String.valueOf(10 + v));
                } catch (Exception ignored) {
                }
            }
            if (el.hasAttr("color")) {
                style.put("color", el.attr("color"));
            }
        }
    }
}
