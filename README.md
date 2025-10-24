# doc-writer

A simple Java library for writing data to Excel (XLSX) and CSV files with no external dependencies.

## Features

### Excel Writer (XLSX)
- Create Excel files with multiple data types: String, Number, Date, DateTime, Time, Currency (USD, EUR, GBP, JPY), Amount
- Support for different currencies with proper formatting
- Support for date, date/time, and time-only fields
- Support for amounts without currency formatting (when currency is in a separate field)
- No external dependencies - uses built-in Java libraries to create XLSX format
- Simple API for adding rows and cells

### CSV Writer
- Configurable separator (comma by default)
- Properly handles data containing separators by enclosing in quotes
- Properly escapes quotes in data by doubling them
- Handles newlines in data

## Building

### Using Maven (Recommended)

Build the project:
```bash
mvn clean compile
```

Run tests:
```bash
mvn test
```

Build package (JAR file):
```bash
mvn package
```

The JAR file will be created in the `target` directory as `doc-writer-1.0.0.jar`.


## Running Tests

### Using Maven (Recommended)

Run all tests:
```bash
mvn test
```


## Running the Demo

Compile and run the demo application:
```bash
javac -d target src/main/java/**.*.java
java -cp target/doc-writer-1.0.0.jar com.docwriter.Demo
```

This will create sample files: `demo_employees.xlsx`, `demo_products.csv`, `demo_special.csv`, and `demo_semicolon.csv`

## Usage Examples

### Excel Writer

```java
import com.docwriter.ExcelWriter;
import java.io.File;
import java.util.Date;

// Create a new Excel writer
ExcelWriter writer = new ExcelWriter("MySheet");

// Add header row
writer.addRow(
    ExcelWriter.createStringCell("Name"),
    ExcelWriter.createStringCell("Age"),
    ExcelWriter.createStringCell("Salary"),
    ExcelWriter.createStringCell("Start Date")
);

// Add data rows
writer.addRow(
    ExcelWriter.createStringCell("John Doe"),
    ExcelWriter.createNumberCell(30),
    ExcelWriter.createCurrencyCell(50000.50),
    ExcelWriter.createDateCell(new Date())
);

// Write to file
writer.writeToFile(new File("output.xlsx"));
```

#### Using new data types

```java
import com.docwriter.ExcelWriter;
import java.io.File;
import java.util.Date;

ExcelWriter writer = new ExcelWriter("ExtendedTypes");

// Add header
writer.addRow(
    ExcelWriter.createStringCell("Product"),
    ExcelWriter.createStringCell("Price (EUR)"),
    ExcelWriter.createStringCell("Price (GBP)"),
    ExcelWriter.createStringCell("Price (JPY)"),
    ExcelWriter.createStringCell("Amount"),
    ExcelWriter.createStringCell("Currency"),
    ExcelWriter.createStringCell("Updated"),
    ExcelWriter.createStringCell("Time")
);

// Add data with different currencies and date/time types
Date now = new Date();
writer.addRow(
    ExcelWriter.createStringCell("Product A"),
    ExcelWriter.createCurrencyEurCell(1200.50),     // EUR: €1,200.50
    ExcelWriter.createCurrencyGbpCell(1050.75),     // GBP: £1,050.75
    ExcelWriter.createCurrencyJpyCell(180000),      // JPY: ¥180,000
    ExcelWriter.createAmountCell(1300.25),          // Amount: 1,300.25 (no currency symbol)
    ExcelWriter.createStringCell("USD"),            // Currency in separate field
    ExcelWriter.createDateTimeCell(now),            // Date and time
    ExcelWriter.createTimeCell(now)                 // Time only
);

writer.writeToFile(new File("output.xlsx"));
```

### CSV Writer

```java
import com.docwriter.CsvWriter;
import java.io.File;

// Create a CSV writer with default comma separator
CsvWriter writer = new CsvWriter();

// Add rows
writer.addRow("Name", "Age", "City");
writer.addRow("John Doe", "30", "New York");
writer.addRow("Jane Smith", "25", "Los Angeles");

// Write to file
writer.writeToFile(new File("output.csv"));

// Use custom separator (e.g., semicolon)
CsvWriter semicolonWriter = new CsvWriter(';');
semicolonWriter.addRow("Field1", "Field2");
semicolonWriter.addRow("Value1", "Value2");
semicolonWriter.writeToFile(new File("output.csv"));
```

## Data Type Support

### Excel Writer
- **String**: Text data with XML escaping
- **Number**: Numeric values (double)
- **Date**: Date values formatted as yyyy-MM-dd
- **DateTime**: Date and time values formatted as yyyy-MM-dd HH:mm:ss
- **Time**: Time-only values formatted as HH:mm:ss
- **Currency (USD)**: Numeric values with USD currency formatting ($#,##0.00)
- **Currency (EUR)**: Numeric values with EUR currency formatting (€#,##0.00)
- **Currency (GBP)**: Numeric values with GBP currency formatting (£#,##0.00)
- **Currency (JPY)**: Numeric values with JPY currency formatting (¥#,##0)
- **Amount**: Numeric values formatted with thousand separators and two decimal places (#,##0.00), without currency symbol. Use this when the currency is stored in a separate field.

### CSV Writer
- All data is treated as strings
- Automatic quoting when data contains:
  - The separator character
  - Quote characters
  - Newline characters
- Quote escaping by doubling quotes (CSV standard)

## Requirements

- Java 8 or higher
- No external dependencies
