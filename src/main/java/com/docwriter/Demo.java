package com.docwriter;

import java.io.File;
import java.util.Date;
import java.util.Calendar;

/**
 * Demo application showing how to use ExcelWriter and CsvWriter
 */
public class Demo {
    
    public static void main(String[] args) {
        System.out.println("Doc-Writer Demo Application");
        System.out.println("===========================\n");
        
        demoExcelWriter();
        System.out.println();
        demoNewDataTypes();
        System.out.println();
        demoCsvWriter();
    }
    
    private static void demoExcelWriter() {
        System.out.println("Excel Writer Demo:");
        System.out.println("------------------");
        
        try {
            // Create a new Excel writer with sheet name
            ExcelWriter writer = new ExcelWriter("EmployeeData");
            
            // Add header row
            writer.addRow(
                ExcelWriter.createStringCell("Employee Name"),
                ExcelWriter.createStringCell("Age"),
                ExcelWriter.createStringCell("Department"),
                ExcelWriter.createStringCell("Salary"),
                ExcelWriter.createStringCell("Hire Date")
            );
            
            // Add sample data
            writer.addRow(
                ExcelWriter.createStringCell("John Doe"),
                ExcelWriter.createNumberCell(30),
                ExcelWriter.createStringCell("Engineering"),
                ExcelWriter.createCurrencyCell(75000.00),
                ExcelWriter.createDateCell(new Date(122, 0, 15)) // Jan 15, 2022
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("Jane Smith"),
                ExcelWriter.createNumberCell(28),
                ExcelWriter.createStringCell("Marketing"),
                ExcelWriter.createCurrencyCell(65000.50),
                ExcelWriter.createDateCell(new Date(122, 5, 20)) // Jun 20, 2022
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("Bob <>&\" Special"),
                ExcelWriter.createNumberCell(35),
                ExcelWriter.createStringCell("Sales"),
                ExcelWriter.createCurrencyCell(80000.75),
                ExcelWriter.createDateCell(new Date())
            );
            
            // Write to file
            File outputFile = new File("demo_employees.xlsx");
            writer.writeToFile(outputFile);
            
            System.out.println("✓ Created Excel file: " + outputFile.getAbsolutePath());
            System.out.println("  File size: " + outputFile.length() + " bytes");
            
        } catch (Exception e) {
            System.err.println("Error creating Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static void demoNewDataTypes() {
        System.out.println("New Data Types Demo:");
        System.out.println("--------------------");
        
        try {
            // Create a new Excel writer
            ExcelWriter writer = new ExcelWriter("ExtendedDataTypes");
            
            // Add header row
            writer.addRow(
                ExcelWriter.createStringCell("Product"),
                ExcelWriter.createStringCell("Price (EUR)"),
                ExcelWriter.createStringCell("Price (GBP)"),
                ExcelWriter.createStringCell("Price (JPY)"),
                ExcelWriter.createStringCell("Amount"),
                ExcelWriter.createStringCell("Currency"),
                ExcelWriter.createStringCell("Last Updated"),
                ExcelWriter.createStringCell("Update Time")
            );
            
            // Create specific date and time
            Calendar cal = Calendar.getInstance();
            cal.set(2024, Calendar.OCTOBER, 24, 14, 30, 45);
            Date testDateTime = cal.getTime();
            
            // Add sample data with new types
            writer.addRow(
                ExcelWriter.createStringCell("Product A"),
                ExcelWriter.createCurrencyEurCell(1200.50),
                ExcelWriter.createCurrencyGbpCell(1050.75),
                ExcelWriter.createCurrencyJpyCell(180000),
                ExcelWriter.createAmountCell(1300.25),
                ExcelWriter.createStringCell("USD"),
                ExcelWriter.createDateTimeCell(testDateTime),
                ExcelWriter.createTimeCell(testDateTime)
            );
            
            writer.addRow(
                ExcelWriter.createStringCell("Product B"),
                ExcelWriter.createCurrencyEurCell(850.00),
                ExcelWriter.createCurrencyGbpCell(750.50),
                ExcelWriter.createCurrencyJpyCell(125000),
                ExcelWriter.createAmountCell(920.00),
                ExcelWriter.createStringCell("EUR"),
                ExcelWriter.createDateTimeCell(new Date()),
                ExcelWriter.createTimeCell(new Date())
            );
            
            // Write to file
            File outputFile = new File("demo_new_types.xlsx");
            writer.writeToFile(outputFile);
            
            System.out.println("✓ Created Excel file with new data types: " + outputFile.getAbsolutePath());
            System.out.println("  File size: " + outputFile.length() + " bytes");
            System.out.println("  Features demonstrated:");
            System.out.println("    - Multiple currencies (EUR, GBP, JPY)");
            System.out.println("    - Date/Time fields");
            System.out.println("    - Time-only fields");
            System.out.println("    - Amount without currency (currency in separate field)");
            
        } catch (Exception e) {
            System.err.println("Error creating Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static void demoCsvWriter() {
        System.out.println("CSV Writer Demo:");
        System.out.println("----------------");
        
        try {
            // Demo 1: Basic CSV with comma separator
            CsvWriter writer = new CsvWriter();
            
            writer.addRow("Product", "Price", "Description");
            writer.addRow("Apple", "$2.99", "Fresh red apples");
            writer.addRow("Orange", "$3.49", "Juicy, sweet oranges");
            writer.addRow("Banana", "$1.99", "Yellow bananas");
            
            File csvFile = new File("demo_products.csv");
            writer.writeToFile(csvFile);
            
            System.out.println("✓ Created CSV file: " + csvFile.getAbsolutePath());
            
            // Demo 2: CSV with data containing commas and quotes
            CsvWriter specialWriter = new CsvWriter();
            
            specialWriter.addRow("Name", "Description");
            specialWriter.addRow("Product A", "Contains, commas, in description");
            specialWriter.addRow("Product \"B\"", "Has \"quotes\" in name and description");
            specialWriter.addRow("Product C", "Normal description");
            
            File specialCsvFile = new File("demo_special.csv");
            specialWriter.writeToFile(specialCsvFile);
            
            System.out.println("✓ Created CSV with special chars: " + specialCsvFile.getAbsolutePath());
            
            // Demo 3: CSV with custom separator (semicolon)
            CsvWriter semicolonWriter = new CsvWriter(';');
            
            semicolonWriter.addRow("Field1", "Field2", "Field3");
            semicolonWriter.addRow("Value1", "Value2", "Value3");
            semicolonWriter.addRow("Data A", "Data B; with semicolon", "Data C");
            
            File semicolonFile = new File("demo_semicolon.csv");
            semicolonWriter.writeToFile(semicolonFile);
            
            System.out.println("✓ Created CSV with semicolon separator: " + semicolonFile.getAbsolutePath());
            
        } catch (Exception e) {
            System.err.println("Error creating CSV file: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
