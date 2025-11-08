package io.github.fivefish130.html2excel.richtext.handler;

import io.github.fivefish130.html2excel.richtext.config.ConverterConfig;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

/**
 * Handler for downloading and embedding images into Excel cells
 *
 * @author fivefish130
 */
public class ImageHandler {

    private static final Logger log = LoggerFactory.getLogger(ImageHandler.class);

    private final ConverterConfig config;

    public ImageHandler(ConverterConfig config) {
        this.config = config;
    }

    /**
     * Process all images in HTML element and embed into cell
     *
     * @param body HTML element
     * @param cell Target cell
     */
    public void processImages(Element body, XSSFCell cell) {
        if (!config.isEnableImageDownload()) {
            return;
        }

        Elements imgElements = body.select("img");
        if (imgElements.isEmpty()) {
            return;
        }

        XSSFSheet sheet = cell.getSheet();
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        int rowIndex = cell.getRowIndex();
        int colIndex = cell.getColumnIndex();

        int imageIndex = 0;
        for (Element img : imgElements) {
            String src = img.attr("src");
            if (src == null || src.trim().isEmpty()) {
                continue;
            }

            try {
                byte[] imageBytes = downloadImage(src);
                if (imageBytes != null && imageBytes.length > 0) {
                    int pictureType = detectImageType(imageBytes, src);
                    int pictureIdx = cell.getSheet().getWorkbook().addPicture(imageBytes, pictureType);

                    ClientAnchor anchor = drawing.createAnchor(
                            0, 0, 0, 0,
                            colIndex, rowIndex + imageIndex,
                            colIndex + 1, rowIndex + imageIndex + 1
                    );

                    anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
                    XSSFPicture picture = drawing.createPicture(anchor, pictureIdx);

                    log.info("Successfully embedded image from {} into cell [{}, {}]",
                            src, rowIndex, colIndex);
                    imageIndex++;
                }
            } catch (Exception e) {
                log.warn("Failed to download or embed image from {}: {}", src, e.getMessage());
            }
        }
    }

    /**
     * Download image from URL
     *
     * @param imageUrl Image URL
     * @return Image bytes, or null if download fails
     */
    private byte[] downloadImage(String imageUrl) {
        InputStream inputStream = null;
        BufferedInputStream bufferedInputStream = null;
        ByteArrayOutputStream outputStream = null;

        try {
            URL url = new URL(imageUrl);
            URLConnection connection = url.openConnection();
            connection.setConnectTimeout(config.getImageConnectTimeout());
            connection.setReadTimeout(config.getImageReadTimeout());
            connection.setRequestProperty("User-Agent",
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");

            inputStream = connection.getInputStream();
            bufferedInputStream = new BufferedInputStream(inputStream);
            outputStream = new ByteArrayOutputStream();

            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = bufferedInputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }

            return outputStream.toByteArray();

        } catch (Exception e) {
            log.warn("Failed to download image from {}: {}", imageUrl, e.getMessage());
            return null;
        } finally {
            try {
                if (outputStream != null) outputStream.close();
                if (bufferedInputStream != null) bufferedInputStream.close();
                if (inputStream != null) inputStream.close();
            } catch (Exception ignored) {
            }
        }
    }

    /**
     * Detect image type from bytes or filename
     *
     * @param imageBytes Image bytes
     * @param filename Filename
     * @return POI picture type constant
     */
    private int detectImageType(byte[] imageBytes, String filename) {
        // Check by magic number
        if (imageBytes.length >= 4) {
            // PNG: 89 50 4E 47
            if (imageBytes[0] == (byte) 0x89 && imageBytes[1] == 0x50 &&
                    imageBytes[2] == 0x4E && imageBytes[3] == 0x47) {
                return Workbook.PICTURE_TYPE_PNG;
            }

            // JPEG: FF D8 FF
            if (imageBytes[0] == (byte) 0xFF && imageBytes[1] == (byte) 0xD8 &&
                    imageBytes[2] == (byte) 0xFF) {
                return Workbook.PICTURE_TYPE_JPEG;
            }

            // GIF: 47 49 46
            if (imageBytes[0] == 0x47 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46) {
                log.warn("GIF format not fully supported, treating as JPEG");
                return Workbook.PICTURE_TYPE_JPEG;
            }

            // BMP: 42 4D
            if (imageBytes[0] == 0x42 && imageBytes[1] == 0x4D) {
                log.warn("BMP format not fully supported, treating as PNG");
                return Workbook.PICTURE_TYPE_PNG;
            }
        }

        // Check by filename extension
        String lowerFilename = filename.toLowerCase();
        if (lowerFilename.endsWith(".png")) {
            return Workbook.PICTURE_TYPE_PNG;
        } else if (lowerFilename.endsWith(".jpg") || lowerFilename.endsWith(".jpeg")) {
            return Workbook.PICTURE_TYPE_JPEG;
        } else if (lowerFilename.endsWith(".gif")) {
            return Workbook.PICTURE_TYPE_JPEG;
        } else if (lowerFilename.endsWith(".bmp")) {
            return Workbook.PICTURE_TYPE_PNG;
        }

        // Default to JPEG
        return Workbook.PICTURE_TYPE_JPEG;
    }
}
