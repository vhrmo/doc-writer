package com.docwriter;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Simple CSV writer with configurable separator.
 * Properly handles data containing separators and quotes.
 */
public class CsvWriter {
    private char separator;
    private List<String[]> rows;
    
    /**
     * Creates a new CSV writer with comma as default separator
     */
    public CsvWriter() {
        this(',');
    }
    
    /**
     * Creates a new CSV writer with specified separator
     * @param separator Character to use as field separator
     */
    public CsvWriter(char separator) {
        this.separator = separator;
        this.rows = new ArrayList<>();
    }
    
    /**
     * Adds a row of data to the CSV
     * @param rowData Array of string values
     */
    public void addRow(String... rowData) {
        rows.add(rowData);
    }
    
    /**
     * Writes the CSV data to the specified writer
     * @param writer Writer to write to
     * @throws IOException If writing fails
     */
    public void write(Writer writer) throws IOException {
        for (String[] row : rows) {
            for (int i = 0; i < row.length; i++) {
                if (i > 0) {
                    writer.write(separator);
                }
                writer.write(escapeField(row[i]));
            }
            writer.write("\n");
        }
        writer.flush();
    }
    
    /**
     * Writes the CSV data to a file
     * @param file Output file
     * @throws IOException If writing fails
     */
    public void writeToFile(File file) throws IOException {
        try (FileWriter fw = new FileWriter(file)) {
            write(fw);
        }
    }
    
    /**
     * Writes the CSV data to an output stream
     * @param outputStream Output stream to write to
     * @throws IOException If writing fails
     */
    public void write(OutputStream outputStream) throws IOException {
        try (OutputStreamWriter osw = new OutputStreamWriter(outputStream, "UTF-8")) {
            write(osw);
        }
    }
    
    /**
     * Escapes a field value according to CSV rules:
     * - If the field contains the separator, quotes, or newlines, it must be enclosed in quotes
     * - Quotes within the field must be escaped by doubling them
     * @param field The field value to escape
     * @return The escaped field value
     */
    private String escapeField(String field) {
        if (field == null) {
            return "";
        }
        
        boolean needsQuotes = false;
        
        // Check if field contains separator, quotes, or newlines
        if (field.indexOf(separator) >= 0 || 
            field.indexOf('"') >= 0 || 
            field.indexOf('\n') >= 0 || 
            field.indexOf('\r') >= 0) {
            needsQuotes = true;
        }
        
        // Escape quotes by doubling them
        String escaped = field.replace("\"", "\"\"");
        
        // Enclose in quotes if necessary
        if (needsQuotes) {
            return "\"" + escaped + "\"";
        }
        
        return escaped;
    }
    
    /**
     * Gets the current separator character
     * @return The separator character
     */
    public char getSeparator() {
        return separator;
    }
    
    /**
     * Sets the separator character
     * @param separator The new separator character
     */
    public void setSeparator(char separator) {
        this.separator = separator;
    }
}
