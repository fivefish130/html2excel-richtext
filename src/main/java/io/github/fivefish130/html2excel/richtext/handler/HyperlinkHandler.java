package io.github.fivefish130.html2excel.richtext.handler;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 * Handler for hyperlinks in HTML
 *
 * @author fivefish130
 */
public class HyperlinkHandler {

    private final XSSFWorkbook workbook;

    public HyperlinkHandler(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Find first href in HTML element
     *
     * @param element HTML element
     * @return First href URL, or null if not found
     */
    public String findFirstHref(Element element) {
        Elements links = element.select("a[href]");
        if (!links.isEmpty()) {
            return links.get(0).attr("href");
        }
        return null;
    }

    /**
     * Apply hyperlink to cell
     *
     * @param cell Target cell
     * @param href URL string
     */
    public void applyHyperlink(XSSFCell cell, String href) {
        if (href == null || href.trim().isEmpty()) {
            return;
        }

        try {
            XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
            link.setAddress(href);
            cell.setHyperlink(link);
        } catch (Exception ignored) {
            // Ignore hyperlink errors
        }
    }
}
