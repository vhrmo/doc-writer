package com.docwriter;

import java.io.*;
import java.util.Date;

/**
 * Test class for ExcelWriter
 */
public class ExcelWriterTest {
    
    public static void main(String[] args) {
        testBasicExcelCreation();
        testAllDataTypes();
        System.out.println("All ExcelWriter tests passed!");
    }
    
    private static void testBasicExcelCreation() {
        try {
            ExcelWriter writer = new ExcelWriter("TestSheet");
            
            // Add header row
            writer.addRow(
                ExcelWriter.createStringCell("Name"),
                ExcelWriter.createStringCell("Age"),
                ExcelWriter.createStringCell("Salary")
            );
            
            // Add data rows
            writer.addRow(
                ExcelWriter.createStringCell("John Doe"),
                ExcelWriter.createNumberCell(30),
                ExcelWriter.createCurrencyCell(50000.50)
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("Jane Smith"),
                ExcelWriter.createNumberCell(25),
                ExcelWriter.createCurrencyCell(45000.75)
            );
            
            File outputFile = new File("/tmp/test_basic.xlsx");
            writer.writeToFile(outputFile);
            
            // Verify file was created
            if (!outputFile.exists()) {
                throw new RuntimeException("File was not created");
            }
            
            // Verify file is not empty
            if (outputFile.length() == 0) {
                throw new RuntimeException("File is empty");
            }
            
            System.out.println("✓ Basic Excel creation test passed");
            System.out.println("  Created file: " + outputFile.getAbsolutePath() + " (" + outputFile.length() + " bytes)");
            
        } catch (Exception e) {
            System.err.println("✗ Basic Excel creation test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void testAllDataTypes() {
        try {
            ExcelWriter writer = new ExcelWriter();
            
            // Add header
            writer.addRow(
                ExcelWriter.createStringCell("String"),
                ExcelWriter.createStringCell("Number"),
                ExcelWriter.createStringCell("Date"),
                ExcelWriter.createStringCell("Currency")
            );
            
            // Add data with all types
            Date testDate = new Date();
            writer.addRow(
                ExcelWriter.createStringCell("Test <>&\"'"),  // Test XML escaping
                ExcelWriter.createNumberCell(123.456),
                ExcelWriter.createDateCell(testDate),
                ExcelWriter.createCurrencyCell(9999.99)
            );
            
            File outputFile = new File("/tmp/test_datatypes.xlsx");
            writer.writeToFile(outputFile);
            
            if (!outputFile.exists() || outputFile.length() == 0) {
                throw new RuntimeException("File creation failed");
            }
            
            System.out.println("✓ All data types test passed");
            System.out.println("  Created file: " + outputFile.getAbsolutePath() + " (" + outputFile.length() + " bytes)");
            
        } catch (Exception e) {
            System.err.println("✗ All data types test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}
