package io.github.fivefish130.html2excel.richtext.jxls;

import io.github.fivefish130.html2excel.richtext.HtmlToExcelConverter;
import io.github.fivefish130.html2excel.richtext.config.ConverterConfig;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * JXLS custom command for converting HTML to Excel rich text
 *
 * <p>Usage in JXLS template:</p>
 * <pre>
 * jx:html(lastCell="A1" value="product.description")
 * jx:html(lastCell="B2" value="item.htmlContent" enableImageDownload="true")
 * </pre>
 *
 * @author fivefish130
 */
public class HtmlCommand extends AbstractCommand {

    private static final Logger log = LoggerFactory.getLogger(HtmlCommand.class);

    private String value;  // HTML content expression
    private Boolean enableImageDownload = false;
    private Integer connectTimeout;
    private Integer readTimeout;
    private Integer maxCellLength;

    private Area area;

    @Override
    public String getName() {
        return "html";
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        if (value == null || value.trim().isEmpty()) {
            log.warn("HTML value expression is empty for cell: {}", cellRef);
            return Size.ZERO_SIZE;
        }

        try {
            // Evaluate HTML content from context
            Object result = getTransformationConfig()
                    .getExpressionEvaluator()
                    .evaluate(value, context.toMap());

            if (result == null) {
                log.debug("HTML expression '{}' evaluated to null for cell: {}", value, cellRef);
                return Size.ZERO_SIZE;
            }

            String htmlContent = result.toString();

            // Get the target cell from transformer
            org.jxls.transform.poi.PoiTransformer poiTransformer =
                    (org.jxls.transform.poi.PoiTransformer) getTransformer();

            Cell cell = poiTransformer.getWorkbook()
                    .getSheet(cellRef.getSheetName())
                    .getRow(cellRef.getRow())
                    .getCell(cellRef.getCol());

            if (cell == null) {
                cell = poiTransformer.getWorkbook()
                        .getSheet(cellRef.getSheetName())
                        .getRow(cellRef.getRow())
                        .createCell(cellRef.getCol());
            }

            if (!(cell instanceof XSSFCell)) {
                log.error("Cell at {} is not an XSSF cell, HTML conversion skipped", cellRef);
                return Size.ZERO_SIZE;
            }

            XSSFCell xssfCell = (XSSFCell) cell;
            XSSFWorkbook workbook = xssfCell.getSheet().getWorkbook();

            // Create converter with config
            ConverterConfig.Builder configBuilder = ConverterConfig.builder();

            if (enableImageDownload != null && enableImageDownload) {
                configBuilder.enableImageDownload(true);
            }

            if (connectTimeout != null && readTimeout != null) {
                configBuilder.imageTimeout(connectTimeout, readTimeout);
            }

            if (maxCellLength != null) {
                configBuilder.maxCellLength(maxCellLength);
            }

            HtmlToExcelConverter converter = new HtmlToExcelConverter(workbook, configBuilder.build());

            // Apply HTML to cell
            converter.applyHtmlToCell(xssfCell, htmlContent);

            log.debug("Applied HTML to cell: {}", cellRef);

            return new Size(1, 1);

        } catch (Exception e) {
            log.error("Failed to apply HTML to cell: {}", cellRef, e);
            return Size.ZERO_SIZE;
        }
    }

    @Override
    public Command addArea(Area area) {
        if (super.getAreaList().size() >= 1) {
            throw new IllegalArgumentException("HtmlCommand can have only one area");
        }
        this.area = area;
        return super.addArea(area);
    }

    // Getters and setters

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public Boolean getEnableImageDownload() {
        return enableImageDownload;
    }

    public void setEnableImageDownload(Boolean enableImageDownload) {
        this.enableImageDownload = enableImageDownload;
    }

    public Integer getConnectTimeout() {
        return connectTimeout;
    }

    public void setConnectTimeout(Integer connectTimeout) {
        this.connectTimeout = connectTimeout;
    }

    public Integer getReadTimeout() {
        return readTimeout;
    }

    public void setReadTimeout(Integer readTimeout) {
        this.readTimeout = readTimeout;
    }

    public Integer getMaxCellLength() {
        return maxCellLength;
    }

    public void setMaxCellLength(Integer maxCellLength) {
        this.maxCellLength = maxCellLength;
    }
}
