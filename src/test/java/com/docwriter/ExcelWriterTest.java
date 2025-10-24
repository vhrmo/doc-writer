package com.docwriter;

import java.io.*;
import java.util.Date;
import java.util.Calendar;

/**
 * Test class for ExcelWriter
 */
public class ExcelWriterTest {
    
    public static void main(String[] args) {
        testBasicExcelCreation();
        testAllDataTypes();
        testNewCurrencyTypes();
        testDateTimeAndTime();
        testAmountWithoutCurrency();
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
    
    private static void testNewCurrencyTypes() {
        try {
            ExcelWriter writer = new ExcelWriter("CurrencyTest");
            
            // Add header
            writer.addRow(
                ExcelWriter.createStringCell("Currency"),
                ExcelWriter.createStringCell("Amount")
            );
            
            // Test different currencies
            writer.addRow(
                ExcelWriter.createStringCell("USD"),
                ExcelWriter.createCurrencyCell(1000.50)
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("EUR"),
                ExcelWriter.createCurrencyEurCell(850.75)
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("GBP"),
                ExcelWriter.createCurrencyGbpCell(750.25)
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("JPY"),
                ExcelWriter.createCurrencyJpyCell(120000)
            );
            
            File outputFile = new File("/tmp/test_currencies.xlsx");
            writer.writeToFile(outputFile);
            
            if (!outputFile.exists() || outputFile.length() == 0) {
                throw new RuntimeException("File creation failed");
            }
            
            System.out.println("✓ New currency types test passed");
            System.out.println("  Created file: " + outputFile.getAbsolutePath() + " (" + outputFile.length() + " bytes)");
            
        } catch (Exception e) {
            System.err.println("✗ New currency types test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void testDateTimeAndTime() {
        try {
            ExcelWriter writer = new ExcelWriter("DateTimeTest");
            
            // Add header
            writer.addRow(
                ExcelWriter.createStringCell("Type"),
                ExcelWriter.createStringCell("Value")
            );
            
            // Create specific date and time for testing
            Calendar cal = Calendar.getInstance();
            cal.set(2024, Calendar.JANUARY, 15, 14, 30, 45);
            Date testDateTime = cal.getTime();
            
            // Test date, datetime, and time
            writer.addRow(
                ExcelWriter.createStringCell("Date"),
                ExcelWriter.createDateCell(testDateTime)
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("DateTime"),
                ExcelWriter.createDateTimeCell(testDateTime)
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("Time"),
                ExcelWriter.createTimeCell(testDateTime)
            );
            
            File outputFile = new File("/tmp/test_datetime.xlsx");
            writer.writeToFile(outputFile);
            
            if (!outputFile.exists() || outputFile.length() == 0) {
                throw new RuntimeException("File creation failed");
            }
            
            System.out.println("✓ Date/Time types test passed");
            System.out.println("  Created file: " + outputFile.getAbsolutePath() + " (" + outputFile.length() + " bytes)");
            
        } catch (Exception e) {
            System.err.println("✗ Date/Time types test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void testAmountWithoutCurrency() {
        try {
            ExcelWriter writer = new ExcelWriter("AmountTest");
            
            // Add header
            writer.addRow(
                ExcelWriter.createStringCell("Product"),
                ExcelWriter.createStringCell("Amount"),
                ExcelWriter.createStringCell("Currency")
            );
            
            // Test amount without currency formatting
            writer.addRow(
                ExcelWriter.createStringCell("Product A"),
                ExcelWriter.createAmountCell(1234.56),
                ExcelWriter.createStringCell("USD")
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("Product B"),
                ExcelWriter.createAmountCell(9876.54),
                ExcelWriter.createStringCell("EUR")
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("Product C"),
                ExcelWriter.createAmountCell(5432.10),
                ExcelWriter.createStringCell("GBP")
            );
            
            File outputFile = new File("/tmp/test_amount.xlsx");
            writer.writeToFile(outputFile);
            
            if (!outputFile.exists() || outputFile.length() == 0) {
                throw new RuntimeException("File creation failed");
            }
            
            System.out.println("✓ Amount without currency test passed");
            System.out.println("  Created file: " + outputFile.getAbsolutePath() + " (" + outputFile.length() + " bytes)");
            
        } catch (Exception e) {
            System.err.println("✗ Amount without currency test failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}
