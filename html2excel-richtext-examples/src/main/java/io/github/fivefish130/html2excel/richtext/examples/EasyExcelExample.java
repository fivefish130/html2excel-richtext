package io.github.fivefish130.html2excel.richtext.examples;

import com.alibaba.excel.EasyExcel;
import io.github.fivefish130.html2excel.richtext.easyexcel.HtmlCell;
import io.github.fivefish130.html2excel.richtext.easyexcel.HtmlCellWriteHandler;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Examples of using EasyExcel integration
 *
 * @author fivefish130
 */
public class EasyExcelExample {

    public static void main(String[] args) {
        // Example 1: Basic HTML annotation
        basicExample();

        // Example 2: Multiple HTML fields
        multipleFieldsExample();

        // Example 3: HTML with images
        imageExample();

        System.out.println("All EasyExcel examples completed!");
    }

    /**
     * Example 1: Basic usage with @HtmlCell annotation
     */
    public static void basicExample() {
        String fileName = "easyexcel-basic.xlsx";

        List<SimpleProduct> data = new ArrayList<>();

        SimpleProduct p1 = new SimpleProduct();
        p1.setName("Laptop");
        p1.setDescription("<p><b>High Performance</b> laptop with <span style='color:red'>amazing</span> features</p>");
        data.add(p1);

        SimpleProduct p2 = new SimpleProduct();
        p2.setName("Mouse");
        p2.setDescription("<p><i>Ergonomic</i> design with <u>wireless</u> connectivity</p>");
        data.add(p2);

        EasyExcel.write(fileName, SimpleProduct.class)
            .registerWriteHandler(new HtmlCellWriteHandler())
            .sheet("Products")
            .doWrite(data);

        System.out.println("✓ Example 1: Basic EasyExcel -> " + fileName);
    }

    /**
     * Example 2: Multiple HTML fields with different configurations
     */
    public static void multipleFieldsExample() {
        String fileName = "easyexcel-multiple.xlsx";

        List<DetailedProduct> data = new ArrayList<>();

        DetailedProduct p1 = new DetailedProduct();
        p1.setName("Smartphone");
        p1.setShortDescription("<p><b>Latest model</b></p>");
        p1.setFullDescription(
            "<p><b>Features:</b></p>" +
            "<ul>" +
            "  <li>6.5\" OLED display</li>" +
            "  <li><span style='color:blue'>5G connectivity</span></li>" +
            "  <li>Triple camera</li>" +
            "</ul>");
        p1.setSpecifications(
            "<table>" +
            "  <tr><td><b>CPU</b></td><td>Snapdragon 888</td></tr>" +
            "  <tr><td><b>RAM</b></td><td>8GB</td></tr>" +
            "  <tr><td><b>Storage</b></td><td>256GB</td></tr>" +
            "</table>");
        data.add(p1);

        DetailedProduct p2 = new DetailedProduct();
        p2.setName("Tablet");
        p2.setShortDescription("<p><i>Portable</i> and <u>powerful</u></p>");
        p2.setFullDescription(
            "<p style='background-color:#FFFF00'>Perfect for work and entertainment</p>");
        p2.setSpecifications(
            "<table>" +
            "  <tr><td><b>Screen</b></td><td>10.2\"</td></tr>" +
            "  <tr><td><b>Battery</b></td><td>10 hours</td></tr>" +
            "</table>");
        data.add(p2);

        EasyExcel.write(fileName, DetailedProduct.class)
            .registerWriteHandler(new HtmlCellWriteHandler())
            .sheet("Detailed Products")
            .doWrite(data);

        System.out.println("✓ Example 2: Multiple fields -> " + fileName);
    }

    /**
     * Example 3: HTML with image download enabled
     */
    public static void imageExample() {
        String fileName = "easyexcel-images.xlsx";

        List<ProductWithImage> data = new ArrayList<>();

        ProductWithImage p1 = new ProductWithImage();
        p1.setName("Camera");
        p1.setDescription(
            "<p><b>Professional camera</b></p>" +
            "<p>Image: <img src='https://via.placeholder.com/100x100'/></p>");
        data.add(p1);

        // Note: Image download requires internet connection
        EasyExcel.write(fileName, ProductWithImage.class)
            .registerWriteHandler(new HtmlCellWriteHandler())
            .sheet("Products with Images")
            .doWrite(data);

        System.out.println("✓ Example 3: With images (if online) -> " + fileName);
    }

    /**
     * Simple product with one HTML field
     */
    public static class SimpleProduct {
        private String name;

        @HtmlCell
        private String description;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getDescription() {
            return description;
        }

        public void setDescription(String description) {
            this.description = description;
        }
    }

    /**
     * Detailed product with multiple HTML fields
     */
    public static class DetailedProduct {
        private String name;

        @HtmlCell
        private String shortDescription;

        @HtmlCell
        private String fullDescription;

        @HtmlCell
        private String specifications;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getShortDescription() {
            return shortDescription;
        }

        public void setShortDescription(String shortDescription) {
            this.shortDescription = shortDescription;
        }

        public String getFullDescription() {
            return fullDescription;
        }

        public void setFullDescription(String fullDescription) {
            this.fullDescription = fullDescription;
        }

        public String getSpecifications() {
            return specifications;
        }

        public void setSpecifications(String specifications) {
            this.specifications = specifications;
        }
    }

    /**
     * Product with HTML field that includes images
     */
    public static class ProductWithImage {
        private String name;

        @HtmlCell(enableImageDownload = true, connectTimeout = 5000, readTimeout = 10000)
        private String description;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getDescription() {
            return description;
        }

        public void setDescription(String description) {
            this.description = description;
        }
    }

    /**
     * Advanced example with custom truncation
     */
    @SuppressWarnings("unused")
    public static class AdvancedProduct {
        private String name;

        @HtmlCell(maxCellLength = 1000, truncateSuffix = "... (truncated)")
        private String description;

        // Getters and setters omitted for brevity
    }
}
