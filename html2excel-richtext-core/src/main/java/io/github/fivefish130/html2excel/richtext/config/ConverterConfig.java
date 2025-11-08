package io.github.fivefish130.html2excel.richtext.config;

/**
 * Configuration for HtmlToExcelConverter
 *
 * @author fivefish130
 */
public class ConverterConfig {

    // Image download settings
    private boolean enableImageDownload = true;
    private int imageConnectTimeout = 3000;  // 3 seconds
    private int imageReadTimeout = 10000;    // 10 seconds

    // Text processing settings
    private int maxCellLength = 32767;
    private String truncateSuffix = "...(truncated)";

    // Cache settings
    private boolean enableFontCache = true;
    private boolean enableStyleCache = true;

    public ConverterConfig() {
    }

    private ConverterConfig(Builder builder) {
        this.enableImageDownload = builder.enableImageDownload;
        this.imageConnectTimeout = builder.imageConnectTimeout;
        this.imageReadTimeout = builder.imageReadTimeout;
        this.maxCellLength = builder.maxCellLength;
        this.truncateSuffix = builder.truncateSuffix;
        this.enableFontCache = builder.enableFontCache;
        this.enableStyleCache = builder.enableStyleCache;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private boolean enableImageDownload = true;
        private int imageConnectTimeout = 3000;
        private int imageReadTimeout = 10000;
        private int maxCellLength = 32767;
        private String truncateSuffix = "...(truncated)";
        private boolean enableFontCache = true;
        private boolean enableStyleCache = true;

        public Builder enableImageDownload(boolean enable) {
            this.enableImageDownload = enable;
            return this;
        }

        public Builder imageTimeout(int connectTimeout, int readTimeout) {
            this.imageConnectTimeout = connectTimeout;
            this.imageReadTimeout = readTimeout;
            return this;
        }

        public Builder maxCellLength(int maxLength) {
            this.maxCellLength = maxLength;
            return this;
        }

        public Builder truncateSuffix(String suffix) {
            this.truncateSuffix = suffix;
            return this;
        }

        public Builder enableFontCache(boolean enable) {
            this.enableFontCache = enable;
            return this;
        }

        public Builder enableStyleCache(boolean enable) {
            this.enableStyleCache = enable;
            return this;
        }

        public ConverterConfig build() {
            return new ConverterConfig(this);
        }
    }

    // Getters
    public boolean isEnableImageDownload() { return enableImageDownload; }
    public int getImageConnectTimeout() { return imageConnectTimeout; }
    public int getImageReadTimeout() { return imageReadTimeout; }
    public int getMaxCellLength() { return maxCellLength; }
    public String getTruncateSuffix() { return truncateSuffix; }
    public boolean isEnableFontCache() { return enableFontCache; }
    public boolean isEnableStyleCache() { return enableStyleCache; }

    // Setters for non-builder usage
    public void setEnableImageDownload(boolean enable) { this.enableImageDownload = enable; }
    public void setImageConnectTimeout(int timeout) { this.imageConnectTimeout = timeout; }
    public void setImageReadTimeout(int timeout) { this.imageReadTimeout = timeout; }
}
