package io.github.fivefish130.html2excel.richtext.easyexcel;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import io.github.fivefish130.html2excel.richtext.HtmlToExcelConverter;
import io.github.fivefish130.html2excel.richtext.config.ConverterConfig;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * EasyExcel cell write handler for HTML to rich text conversion
 *
 * <p>This handler automatically converts HTML content to Excel rich text for fields
 * annotated with {@link HtmlCell}.</p>
 *
 * <p>Usage example:</p>
 * <pre>
 * EasyExcel.write(outputStream, Product.class)
 *     .registerWriteHandler(new HtmlCellWriteHandler())
 *     .sheet("Products")
 *     .doWrite(dataList);
 * </pre>
 *
 * @author fivefish130
 */
public class HtmlCellWriteHandler implements CellWriteHandler {

    private static final Logger log = LoggerFactory.getLogger(HtmlCellWriteHandler.class);

    private final Map<String, HtmlCellConfig> htmlCellConfigCache = new ConcurrentHashMap<>();
    private HtmlToExcelConverter defaultConverter;

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder,
                                  WriteTableHolder writeTableHolder,
                                  List<WriteCellData<?>> cellDataList,
                                  Cell cell,
                                  Head head,
                                  Integer relativeRowIndex,
                                  Boolean isHead) {
        // Skip header rows
        if (isHead) {
            return;
        }

        // Check if this field has @HtmlCell annotation
        if (head == null || head.getField() == null) {
            return;
        }

        Field field = head.getField();
        HtmlCell htmlCellAnnotation = field.getAnnotation(HtmlCell.class);

        if (htmlCellAnnotation == null) {
            return;
        }

        // Ensure it's an XSSF cell
        if (!(cell instanceof XSSFCell)) {
            log.warn("Cell is not an XSSF cell, HTML conversion skipped for field: {}", field.getName());
            return;
        }

        try {
            XSSFCell xssfCell = (XSSFCell) cell;
            String htmlContent = cell.getStringCellValue();

            if (htmlContent == null || htmlContent.trim().isEmpty()) {
                return;
            }

            XSSFWorkbook workbook = xssfCell.getSheet().getWorkbook();

            // Get or create converter with config from annotation
            String configKey = getConfigKey(htmlCellAnnotation);
            HtmlCellConfig config = htmlCellConfigCache.computeIfAbsent(configKey, k ->
                    new HtmlCellConfig(htmlCellAnnotation));

            HtmlToExcelConverter converter = config.getConverter(workbook);

            // Apply HTML conversion
            converter.applyHtmlToCell(xssfCell, htmlContent);

            log.debug("Converted HTML to rich text for field: {}", field.getName());

        } catch (Exception e) {
            log.error("Failed to convert HTML for field: {}", field.getName(), e);
        }
    }

    private String getConfigKey(HtmlCell annotation) {
        return annotation.enableImageDownload() + "_" +
                annotation.connectTimeout() + "_" +
                annotation.readTimeout() + "_" +
                annotation.maxCellLength() + "_" +
                annotation.truncateSuffix();
    }

    /**
     * Inner class to cache converter configuration
     */
    private static class HtmlCellConfig {
        private final ConverterConfig converterConfig;
        private HtmlToExcelConverter converter;

        public HtmlCellConfig(HtmlCell annotation) {
            ConverterConfig.Builder builder = ConverterConfig.builder();

            if (annotation.enableImageDownload()) {
                builder.enableImageDownload(true)
                        .imageTimeout(annotation.connectTimeout(), annotation.readTimeout());
            }

            builder.maxCellLength(annotation.maxCellLength())
                    .truncateSuffix(annotation.truncateSuffix());

            this.converterConfig = builder.build();
        }

        public synchronized HtmlToExcelConverter getConverter(XSSFWorkbook workbook) {
            if (converter == null) {
                converter = new HtmlToExcelConverter(workbook, converterConfig);
            }
            return converter;
        }
    }
}
