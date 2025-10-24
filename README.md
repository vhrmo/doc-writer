# doc-writer

A simple Java library for writing data to Excel (XLSX) and CSV files with no external dependencies.

## Features

### Excel Writer (XLSX)
- Create Excel files with basic data types: String, Number, Date, Currency
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
- **Currency**: Numeric values with currency formatting ($#,##0.00)

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
