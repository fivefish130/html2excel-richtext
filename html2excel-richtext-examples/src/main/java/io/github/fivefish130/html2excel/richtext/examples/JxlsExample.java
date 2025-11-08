package io.github.fivefish130.html2excel.richtext.examples;

import io.github.fivefish130.html2excel.richtext.jxls.JxlsHtmlHelper;
import org.jxls.common.Context;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Examples of using JXLS integration
 *
 * <p>To run this example, you need to create a template Excel file first.
 * See the instructions in the main() method.</p>
 *
 * @author fivefish130
 */
public class JxlsExample {

    public static void main(String[] args) {
        System.out.println("JXLS Example");
        System.out.println("============");
        System.out.println();
        System.out.println("To use JXLS HTML integration:");
        System.out.println();
        System.out.println("1. Create an Excel template file (e.g., template.xlsx)");
        System.out.println("2. Add a cell comment with JXLS command:");
        System.out.println("   jx:html(lastCell=\"A1\" value=\"product.description\")");
        System.out.println();
        System.out.println("3. Use the code below to process the template:");
        System.out.println();

        printExampleCode();

        System.out.println();
        System.out.println("For a working example, uncomment and run processTemplateExample()");
        System.out.println("after creating the template file.");

        // Uncomment to run actual example (requires template file)
        // processTemplateExample();
    }

    /**
     * Print example code for reference
     */
    private static void printExampleCode() {
        System.out.println("// Example code:");
        System.out.println("// -------------");
        System.out.println();
        System.out.println("// 1. Prepare data");
        System.out.println("Product product = new Product();");
        System.out.println("product.setName(\"Laptop\");");
        System.out.println("product.setDescription(\"<p><b>High Performance</b> laptop with <span style='color:red'>amazing</span> features</p>\");");
        System.out.println();
        System.out.println("// 2. Create context");
        System.out.println("Context context = new Context();");
        System.out.println("context.putVar(\"product\", product);");
        System.out.println();
        System.out.println("// 3. Process template");
        System.out.println("try (InputStream template = new FileInputStream(\"template.xlsx\");");
        System.out.println("     OutputStream output = new FileOutputStream(\"output.xlsx\")) {");
        System.out.println("    JxlsHtmlHelper.processTemplate(template, output, context);");
        System.out.println("}");
    }

    /**
     * Actual working example (requires template file)
     */
    @SuppressWarnings("unused")
    private static void processTemplateExample() {
        try {
            // Prepare sample data
            List<Product> products = new ArrayList<>();

            Product product1 = new Product();
            product1.setName("Laptop");
            product1.setDescription(
                "<p><b>High Performance</b> laptop</p>" +
                "<ul>" +
                "  <li>Intel i7 processor</li>" +
                "  <li><span style='color:blue'>16GB RAM</span></li>" +
                "  <li>512GB SSD</li>" +
                "</ul>");
            products.add(product1);

            Product product2 = new Product();
            product2.setName("Mouse");
            product2.setDescription(
                "<p><i>Ergonomic</i> wireless mouse</p>" +
                "<ul>" +
                "  <li>2.4GHz wireless</li>" +
                "  <li><span style='color:green'>Long battery life</span></li>" +
                "</ul>");
            products.add(product2);

            // Create context
            Context context = new Context();
            context.putVar("products", products);

            // Process template
            try (InputStream template = new FileInputStream("jxls-template.xlsx");
                 OutputStream output = new FileOutputStream("jxls-output.xlsx")) {

                JxlsHtmlHelper.processTemplate(template, output, context);
                System.out.println("✓ JXLS template processed -> jxls-output.xlsx");
            }

        } catch (Exception e) {
            System.err.println("Error processing template: " + e.getMessage());
            System.err.println("Make sure jxls-template.xlsx exists!");
        }
    }

    /**
     * Sample Product class
     */
    public static class Product {
        private String name;
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
     * Example of using custom JXLS configuration
     */
    @SuppressWarnings("unused")
    private static void customConfigExample() {
        try {
            Product product = new Product();
            product.setName("Sample");
            product.setDescription("<p><b>Bold</b> description</p>");

            Context context = new Context();
            context.putVar("product", product);

            JxlsHtmlHelper.ProcessConfig config = JxlsHtmlHelper.ProcessConfig.defaults()
                .useFastFormulaProcessor(false)
                .processFormulas(true)
                .hideTemplateSheet(true)
                .deleteTemplateSheet(false);

            try (InputStream template = new FileInputStream("template.xlsx");
                 OutputStream output = new FileOutputStream("output.xlsx")) {

                JxlsHtmlHelper.processTemplate(template, output, context, config);
            }

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
        }
    }
}
