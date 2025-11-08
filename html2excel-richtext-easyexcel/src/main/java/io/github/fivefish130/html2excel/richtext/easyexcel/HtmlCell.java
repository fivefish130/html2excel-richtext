package io.github.fivefish130.html2excel.richtext.easyexcel;

import java.lang.annotation.*;

/**
 * Annotation to mark fields that contain HTML content and should be converted to Excel rich text
 *
 * <p>Usage example:</p>
 * <pre>
 * public class Product {
 *     private String name;
 *
 *     &#64;HtmlCell
 *     private String description;  // Will be converted from HTML to rich text
 *
 *     &#64;HtmlCell(enableImageDownload = true, connectTimeout = 5000, readTimeout = 10000)
 *     private String detailedInfo;
 * }
 * </pre>
 *
 * @author fivefish130
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface HtmlCell {

    /**
     * Enable image download and embedding from HTML img tags
     *
     * @return true to enable image download
     */
    boolean enableImageDownload() default false;

    /**
     * Connection timeout in milliseconds for image download
     *
     * @return connection timeout (default: 5000ms)
     */
    int connectTimeout() default 5000;

    /**
     * Read timeout in milliseconds for image download
     *
     * @return read timeout (default: 15000ms)
     */
    int readTimeout() default 15000;

    /**
     * Maximum cell content length (Excel limit is 32,767)
     *
     * @return max length (default: 32767)
     */
    int maxCellLength() default 32767;

    /**
     * Truncation suffix when content exceeds max length
     *
     * @return truncation suffix (default: "...")
     */
    String truncateSuffix() default "...";
}
