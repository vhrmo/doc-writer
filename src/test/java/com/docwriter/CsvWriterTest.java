package com.docwriter;

import java.io.*;
import java.nio.file.Files;
import java.util.List;

/**
 * Test class for CsvWriter
 */
public class CsvWriterTest {
    
    public static void main(String[] args) {
        testBasicCsvCreation();
        testQuoteEscaping();
        testSeparatorInData();
        testCustomSeparator();
        testNewlinesInData();
        System.out.println("All CsvWriter tests passed!");
    }
    
    private static void testBasicCsvCreation() {
        try {
            CsvWriter writer = new CsvWriter();
            
            writer.addRow("Name", "Age", "City");
            writer.addRow("John Doe", "30", "New York");
            writer.addRow("Jane Smith", "25", "Los Angeles");
            
            File outputFile = new File("/tmp/test_basic.csv");
            writer.writeToFile(outputFile);
            
            // Verify file was created
            if (!outputFile.exists()) {
                throw new RuntimeException("File was not created");
            }
            
            // Read and verify content
            List<String> lines = Files.readAllLines(outputFile.toPath());
            if (lines.size() != 3) {
                throw new RuntimeException("Expected 3 lines, got " + lines.size());
            }
            
            if (!lines.get(0).equals("Name,Age,City")) {
                throw new RuntimeException("Header line incorrect: " + lines.get(0));
            }
            
            System.out.println("✓ Basic CSV creation test passed");
            System.out.println("  Output: " + lines);
            
        } catch (Exception e) {
            System.err.println("✗ Basic CSV creation test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void testQuoteEscaping() {
        try {
            CsvWriter writer = new CsvWriter();
            
            writer.addRow("Name", "Quote");
            writer.addRow("John \"The Boss\" Doe", "She said \"Hello\"");
            
            File outputFile = new File("/tmp/test_quotes.csv");
            writer.writeToFile(outputFile);
            
            List<String> lines = Files.readAllLines(outputFile.toPath());
            
            // Quotes should be escaped by doubling them and the field should be quoted
            String expectedLine = "\"John \"\"The Boss\"\" Doe\",\"She said \"\"Hello\"\"\"";
            if (!lines.get(1).equals(expectedLine)) {
                throw new RuntimeException("Quote escaping failed. Expected: " + expectedLine + ", Got: " + lines.get(1));
            }
            
            System.out.println("✓ Quote escaping test passed");
            System.out.println("  Output: " + lines.get(1));
            
        } catch (Exception e) {
            System.err.println("✗ Quote escaping test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void testSeparatorInData() {
        try {
            CsvWriter writer = new CsvWriter();
            
            writer.addRow("Product", "Price");
            writer.addRow("Apples, fresh", "$2.99");
            writer.addRow("Oranges, juicy", "$3.49");
            
            File outputFile = new File("/tmp/test_separator.csv");
            writer.writeToFile(outputFile);
            
            List<String> lines = Files.readAllLines(outputFile.toPath());
            
            // Fields containing comma should be quoted
            String expectedLine = "\"Apples, fresh\",$2.99";
            if (!lines.get(1).equals(expectedLine)) {
                throw new RuntimeException("Separator handling failed. Expected: " + expectedLine + ", Got: " + lines.get(1));
            }
            
            System.out.println("✓ Separator in data test passed");
            System.out.println("  Output: " + lines.get(1));
            
        } catch (Exception e) {
            System.err.println("✗ Separator in data test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void testCustomSeparator() {
        try {
            CsvWriter writer = new CsvWriter(';');
            
            writer.addRow("Name", "Description");
            writer.addRow("Product A", "Contains; special chars");
            
            File outputFile = new File("/tmp/test_custom_sep.csv");
            writer.writeToFile(outputFile);
            
            List<String> lines = Files.readAllLines(outputFile.toPath());
            
            // Should use semicolon as separator
            if (!lines.get(0).equals("Name;Description")) {
                throw new RuntimeException("Custom separator not used in header");
            }
            
            // Field containing semicolon should be quoted
            String expectedLine = "Product A;\"Contains; special chars\"";
            if (!lines.get(1).equals(expectedLine)) {
                throw new RuntimeException("Custom separator handling failed. Expected: " + expectedLine + ", Got: " + lines.get(1));
            }
            
            System.out.println("✓ Custom separator test passed");
            System.out.println("  Output: " + lines);
            
        } catch (Exception e) {
            System.err.println("✗ Custom separator test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void testNewlinesInData() {
        try {
            CsvWriter writer = new CsvWriter();
            
            writer.addRow("Field", "Value");
            writer.addRow("Multi\nLine", "Text with\nnewlines");
            
            File outputFile = new File("/tmp/test_newlines.csv");
            writer.writeToFile(outputFile);
            
            String content = new String(Files.readAllBytes(outputFile.toPath()));
            
            // Fields with newlines should be quoted
            if (!content.contains("\"Multi\nLine\"")) {
                throw new RuntimeException("Newline handling failed");
            }
            
            System.out.println("✓ Newlines in data test passed");
            
        } catch (Exception e) {
            System.err.println("✗ Newlines in data test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}
