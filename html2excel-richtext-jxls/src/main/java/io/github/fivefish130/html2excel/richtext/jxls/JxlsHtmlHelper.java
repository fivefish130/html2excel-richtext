package io.github.fivefish130.html2excel.richtext.jxls;

import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Helper class for using HTML conversion with JXLS templates
 *
 * <p>This class simplifies the integration of HTML to Excel rich text conversion
 * with JXLS template processing.</p>
 *
 * <p>To use HTML command in JXLS templates, you need to register it first:</p>
 * <pre>
 * // Register HTML command globally before processing
 * XlsCommentAreaBuilder.addCommandMapping("html", HtmlCommand.class);
 *
 * // Then use in template comments:
 * // jx:html(lastCell="A1" value="product.description")
 * </pre>
 *
 * @author fivefish130
 */
public class JxlsHtmlHelper {

    private static final Logger log = LoggerFactory.getLogger(JxlsHtmlHelper.class);

    static {
        // Register HTML command globally
        XlsCommentAreaBuilder.addCommandMapping("html", HtmlCommand.class);
    }

    /**
     * Process JXLS template with HTML command support
     *
     * @param templateStream JXLS template input stream
     * @param outputStream   Output stream for result Excel
     * @param context        JXLS context with data
     * @throws IOException if template processing fails
     */
    public static void processTemplate(InputStream templateStream,
                                       OutputStream outputStream,
                                       Context context) throws IOException {
        processTemplate(templateStream, outputStream, context, false);
    }

    /**
     * Process JXLS template with HTML command support
     *
     * @param templateStream       JXLS template input stream
     * @param outputStream         Output stream for result Excel
     * @param context              JXLS context with data
     * @param useFastFormulaProcessor Whether to use fast formula processor
     * @throws IOException if template processing fails
     */
    public static void processTemplate(InputStream templateStream,
                                       OutputStream outputStream,
                                       Context context,
                                       boolean useFastFormulaProcessor) throws IOException {
        try {
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();
            jxlsHelper.setUseFastFormulaProcessor(useFastFormulaProcessor);
            jxlsHelper.processTemplate(templateStream, outputStream, context);

            log.debug("JXLS template processed successfully with HTML command support");

        } catch (Exception e) {
            log.error("Failed to process JXLS template", e);
            throw new IOException("Failed to process JXLS template", e);
        }
    }

    /**
     * Process JXLS template with advanced configuration
     *
     * @param templateStream JXLS template input stream
     * @param outputStream   Output stream for result Excel
     * @param context        JXLS context with data
     * @param config         Processing configuration
     * @throws IOException if template processing fails
     */
    public static void processTemplate(InputStream templateStream,
                                       OutputStream outputStream,
                                       Context context,
                                       ProcessConfig config) throws IOException {
        try {
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();

            if (config != null) {
                jxlsHelper.setUseFastFormulaProcessor(config.isUseFastFormulaProcessor());
                jxlsHelper.setProcessFormulas(config.isProcessFormulas());
                jxlsHelper.setHideTemplateSheet(config.isHideTemplateSheet());
                jxlsHelper.setDeleteTemplateSheet(config.isDeleteTemplateSheet());
            }

            jxlsHelper.processTemplate(templateStream, outputStream, context);

            log.debug("JXLS template processed successfully with custom config");

        } catch (Exception e) {
            log.error("Failed to process JXLS template with config", e);
            throw new IOException("Failed to process JXLS template", e);
        }
    }

    /**
     * Configuration for JXLS template processing
     */
    public static class ProcessConfig {
        private boolean useFastFormulaProcessor = false;
        private boolean processFormulas = true;
        private boolean hideTemplateSheet = true;
        private boolean deleteTemplateSheet = false;

        public static ProcessConfig defaults() {
            return new ProcessConfig();
        }

        public ProcessConfig useFastFormulaProcessor(boolean use) {
            this.useFastFormulaProcessor = use;
            return this;
        }

        public ProcessConfig processFormulas(boolean process) {
            this.processFormulas = process;
            return this;
        }

        public ProcessConfig hideTemplateSheet(boolean hide) {
            this.hideTemplateSheet = hide;
            return this;
        }

        public ProcessConfig deleteTemplateSheet(boolean delete) {
            this.deleteTemplateSheet = delete;
            return this;
        }

        // Getters
        public boolean isUseFastFormulaProcessor() {
            return useFastFormulaProcessor;
        }

        public boolean isProcessFormulas() {
            return processFormulas;
        }

        public boolean isHideTemplateSheet() {
            return hideTemplateSheet;
        }

        public boolean isDeleteTemplateSheet() {
            return deleteTemplateSheet;
        }
    }
}
