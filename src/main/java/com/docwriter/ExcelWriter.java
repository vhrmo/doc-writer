package com.docwriter;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Simple Excel (XLSX) writer that creates basic spreadsheet files.
 * Supports string, date, number, and currency data types.
 * No external dependencies required.
 */
public class ExcelWriter {
    private List<List<CellData>> rows;
    private String sheetName;
    private Map<String, Integer> currencyFormats;
    private int nextFormatId = 164;
    private int nextStyleIndex = 2;
    
    /**
     * Cell data structure to hold value and type information
     */
    public static class CellData {
        public enum CellType {
            STRING, NUMBER, DATE, DATETIME, TIME, CURRENCY, AMOUNT
        }
        
        private String value;
        private CellType type;
        private String currencyCode;
        
        public CellData(String value, CellType type) {
            this(value, type, null);
        }
        
        public CellData(String value, CellType type, String currencyCode) {
            this.value = value;
            this.type = type;
            this.currencyCode = currencyCode;
        }
        
        public String getValue() {
            return value;
        }
        
        public CellType getType() {
            return type;
        }
        
        public String getCurrencyCode() {
            return currencyCode;
        }
    }
    
    /**
     * Creates a new Excel writer with default sheet name "Sheet1"
     */
    public ExcelWriter() {
        this("Sheet1");
    }
    
    /**
     * Creates a new Excel writer with specified sheet name
     * @param sheetName Name of the worksheet
     */
    public ExcelWriter(String sheetName) {
        this.sheetName = sheetName;
        this.rows = new ArrayList<>();
        this.currencyFormats = new LinkedHashMap<>();
    }
    
    /**
     * Adds a row of data to the sheet
     * @param rowData Array of cell data
     */
    public void addRow(CellData... rowData) {
        rows.add(Arrays.asList(rowData));
    }
    
    /**
     * Creates a string cell
     * @param value String value
     * @return CellData object
     */
    public static CellData createStringCell(String value) {
        return new CellData(value, CellData.CellType.STRING);
    }
    
    /**
     * Creates a number cell
     * @param value Numeric value
     * @return CellData object
     */
    public static CellData createNumberCell(double value) {
        return new CellData(String.valueOf(value), CellData.CellType.NUMBER);
    }
    
    /**
     * Creates a date cell
     * @param date Date value
     * @return CellData object
     */
    public static CellData createDateCell(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        return new CellData(sdf.format(date), CellData.CellType.DATE);
    }
    
    /**
     * Creates a currency cell (USD by default)
     * @param value Currency value
     * @return CellData object
     */
    public static CellData createCurrencyCell(double value) {
        return createCurrencyCell(value, "USD");
    }
    
    /**
     * Creates a currency cell with specified currency code
     * @param value Currency value
     * @param currencyCode Currency code (e.g., "USD", "EUR", "GBP", "JPY")
     * @return CellData object
     */
    public static CellData createCurrencyCell(double value, String currencyCode) {
        return new CellData(String.valueOf(value), CellData.CellType.CURRENCY, currencyCode);
    }
    
    /**
     * Creates a date/time cell
     * @param date Date/time value
     * @return CellData object
     */
    public static CellData createDateTimeCell(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return new CellData(sdf.format(date), CellData.CellType.DATETIME);
    }
    
    /**
     * Creates a time cell
     * @param date Date value (only time portion will be used)
     * @return CellData object
     */
    public static CellData createTimeCell(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
        return new CellData(sdf.format(date), CellData.CellType.TIME);
    }
    
    /**
     * Creates an amount cell without currency formatting
     * @param value Numeric value
     * @return CellData object
     */
    public static CellData createAmountCell(double value) {
        return new CellData(String.valueOf(value), CellData.CellType.AMOUNT);
    }
    
    /**
     * Gets the currency symbol for a given currency code
     * @param currencyCode Currency code (e.g., "USD", "EUR", "GBP", "JPY")
     * @return Currency symbol
     */
    private String getCurrencySymbol(String currencyCode) {
        if (currencyCode == null) {
            return "$";
        }
        switch (currencyCode.toUpperCase()) {
            case "EUR":
                return "€";
            case "GBP":
                return "£";
            case "JPY":
                return "¥";
            case "USD":
            default:
                return "$";
        }
    }
    
    /**
     * Gets the number format code for a currency
     * @param currencyCode Currency code
     * @return Excel number format code
     */
    private String getCurrencyFormatCode(String currencyCode) {
        String symbol = getCurrencySymbol(currencyCode);
        // JPY typically doesn't use decimals
        if ("JPY".equalsIgnoreCase(currencyCode)) {
            return "&quot;" + symbol + "&quot;#,##0";
        }
        return "&quot;" + symbol + "&quot;#,##0.00";
    }
    
    /**
     * Registers a currency and returns its style index
     * @param currencyCode Currency code
     * @return Style index for the currency
     */
    private int getCurrencyStyleIndex(String currencyCode) {
        if (currencyCode == null) {
            currencyCode = "USD";
        }
        currencyCode = currencyCode.toUpperCase();
        
        if (!currencyFormats.containsKey(currencyCode)) {
            currencyFormats.put(currencyCode, nextStyleIndex++);
        }
        return currencyFormats.get(currencyCode);
    }
    
    /**
     * Writes the Excel file to the specified output stream
     * @param outputStream Output stream to write to
     * @throws IOException If writing fails
     */
    public void write(OutputStream outputStream) throws IOException {
        // First pass: scan for all currencies to build format map
        for (List<CellData> row : rows) {
            for (CellData cell : row) {
                if (cell.getType() == CellData.CellType.CURRENCY) {
                    getCurrencyStyleIndex(cell.getCurrencyCode());
                }
            }
        }
        
        try (ZipOutputStream zipOut = new ZipOutputStream(outputStream)) {
            // Write [Content_Types].xml
            writeContentTypes(zipOut);
            
            // Write _rels/.rels
            writeRels(zipOut);
            
            // Write xl/_rels/workbook.xml.rels
            writeWorkbookRels(zipOut);
            
            // Write xl/workbook.xml
            writeWorkbook(zipOut);
            
            // Write xl/styles.xml
            writeStyles(zipOut);
            
            // Write xl/worksheets/sheet1.xml
            writeSheet(zipOut);
            
            zipOut.finish();
        }
    }
    
    /**
     * Writes the Excel file to a file
     * @param file Output file
     * @throws IOException If writing fails
     */
    public void writeToFile(File file) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            write(fos);
        }
    }
    
    private void writeContentTypes(ZipOutputStream zipOut) throws IOException {
        ZipEntry entry = new ZipEntry("[Content_Types].xml");
        zipOut.putNextEntry(entry);
        
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n" +
                "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n" +
                "</Types>";
        
        zipOut.write(xml.getBytes("UTF-8"));
        zipOut.closeEntry();
    }
    
    private void writeRels(ZipOutputStream zipOut) throws IOException {
        ZipEntry entry = new ZipEntry("_rels/.rels");
        zipOut.putNextEntry(entry);
        
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n" +
                "</Relationships>";
        
        zipOut.write(xml.getBytes("UTF-8"));
        zipOut.closeEntry();
    }
    
    private void writeWorkbookRels(ZipOutputStream zipOut) throws IOException {
        ZipEntry entry = new ZipEntry("xl/_rels/workbook.xml.rels");
        zipOut.putNextEntry(entry);
        
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n" +
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n" +
                "</Relationships>";
        
        zipOut.write(xml.getBytes("UTF-8"));
        zipOut.closeEntry();
    }
    
    private void writeWorkbook(ZipOutputStream zipOut) throws IOException {
        ZipEntry entry = new ZipEntry("xl/workbook.xml");
        zipOut.putNextEntry(entry);
        
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
                "<sheets>\n" +
                "<sheet name=\"" + escapeXml(sheetName) + "\" sheetId=\"1\" r:id=\"rId1\"/>\n" +
                "</sheets>\n" +
                "</workbook>";
        
        zipOut.write(xml.getBytes("UTF-8"));
        zipOut.closeEntry();
    }
    
    private void writeStyles(ZipOutputStream zipOut) throws IOException {
        ZipEntry entry = new ZipEntry("xl/styles.xml");
        zipOut.putNextEntry(entry);
        
        // Build number formats dynamically
        StringBuilder numFmts = new StringBuilder();
        int numFmtCount = 3; // Base formats: date, datetime, time
        
        // Add base formats
        numFmts.append("<numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/>\n");
        numFmts.append("<numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd hh:mm:ss\"/>\n");
        numFmts.append("<numFmt numFmtId=\"166\" formatCode=\"hh:mm:ss\"/>\n");
        
        // Add currency formats
        int formatId = 167;
        for (String currencyCode : currencyFormats.keySet()) {
            numFmts.append("<numFmt numFmtId=\"").append(formatId++).append("\" formatCode=\"");
            numFmts.append(getCurrencyFormatCode(currencyCode)).append("\"/>\n");
            numFmtCount++;
        }
        
        // Add amount format
        numFmts.append("<numFmt numFmtId=\"").append(formatId).append("\" formatCode=\"#,##0.00\"/>\n");
        numFmtCount++;
        
        // Build cell styles
        StringBuilder cellXfs = new StringBuilder();
        int styleCount = 2; // Base styles: default, number
        
        // Add base styles
        cellXfs.append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n"); // 0 - default/string
        cellXfs.append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n"); // 1 - number
        
        // Add date/time styles (indices 2, 3, 4)
        cellXfs.append("<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n"); // 2 - date
        cellXfs.append("<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n"); // 3 - datetime
        cellXfs.append("<xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n"); // 4 - time
        styleCount += 3;
        
        // Add currency styles (starting at index 5)
        formatId = 167;
        for (String currencyCode : currencyFormats.keySet()) {
            cellXfs.append("<xf numFmtId=\"").append(formatId++).append("\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
            styleCount++;
        }
        
        // Add amount style
        cellXfs.append("<xf numFmtId=\"").append(formatId).append("\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
        styleCount++;
        
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
                "<numFmts count=\"" + numFmtCount + "\">\n" +
                numFmts.toString() +
                "</numFmts>\n" +
                "<fonts count=\"1\">\n" +
                "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>\n" +
                "</fonts>\n" +
                "<fills count=\"1\">\n" +
                "<fill><patternFill patternType=\"none\"/></fill>\n" +
                "</fills>\n" +
                "<borders count=\"1\">\n" +
                "<border><left/><right/><top/><bottom/><diagonal/></border>\n" +
                "</borders>\n" +
                "<cellXfs count=\"" + styleCount + "\">\n" +
                cellXfs.toString() +
                "</cellXfs>\n" +
                "</styleSheet>";
        
        zipOut.write(xml.getBytes("UTF-8"));
        zipOut.closeEntry();
    }
    
    private void writeSheet(ZipOutputStream zipOut) throws IOException {
        ZipEntry entry = new ZipEntry("xl/worksheets/sheet1.xml");
        zipOut.putNextEntry(entry);
        
        StringBuilder xml = new StringBuilder();
        xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        xml.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
        xml.append("<sheetData>\n");
        
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            List<CellData> row = rows.get(rowIndex);
            xml.append("<row r=\"").append(rowIndex + 1).append("\">\n");
            
            for (int colIndex = 0; colIndex < row.size(); colIndex++) {
                CellData cell = row.get(colIndex);
                String cellRef = getColumnName(colIndex) + (rowIndex + 1);
                
                if (cell.getType() == CellData.CellType.STRING) {
                    xml.append("<c r=\"").append(cellRef).append("\" t=\"inlineStr\">");
                    xml.append("<is><t>").append(escapeXml(cell.getValue())).append("</t></is>");
                    xml.append("</c>\n");
                } else if (cell.getType() == CellData.CellType.NUMBER) {
                    xml.append("<c r=\"").append(cellRef).append("\" s=\"1\">");
                    xml.append("<v>").append(cell.getValue()).append("</v>");
                    xml.append("</c>\n");
                } else if (cell.getType() == CellData.CellType.DATE) {
                    // Convert date string to Excel serial number
                    xml.append("<c r=\"").append(cellRef).append("\" s=\"2\">");
                    xml.append("<v>").append(dateToExcelSerial(cell.getValue())).append("</v>");
                    xml.append("</c>\n");
                } else if (cell.getType() == CellData.CellType.DATETIME) {
                    // Convert datetime string to Excel serial number
                    xml.append("<c r=\"").append(cellRef).append("\" s=\"3\">");
                    xml.append("<v>").append(dateTimeToExcelSerial(cell.getValue())).append("</v>");
                    xml.append("</c>\n");
                } else if (cell.getType() == CellData.CellType.TIME) {
                    // Convert time string to Excel serial fraction
                    xml.append("<c r=\"").append(cellRef).append("\" s=\"4\">");
                    xml.append("<v>").append(timeToExcelSerial(cell.getValue())).append("</v>");
                    xml.append("</c>\n");
                } else if (cell.getType() == CellData.CellType.CURRENCY) {
                    // Use dynamic currency style index (starts at 5)
                    int styleIndex = 5 + getCurrencyStyleIndex(cell.getCurrencyCode()) - 2;
                    xml.append("<c r=\"").append(cellRef).append("\" s=\"").append(styleIndex).append("\">");
                    xml.append("<v>").append(cell.getValue()).append("</v>");
                    xml.append("</c>\n");
                } else if (cell.getType() == CellData.CellType.AMOUNT) {
                    // Amount style is after all currency styles
                    int amountStyleIndex = 5 + currencyFormats.size();
                    xml.append("<c r=\"").append(cellRef).append("\" s=\"").append(amountStyleIndex).append("\">");
                    xml.append("<v>").append(cell.getValue()).append("</v>");
                    xml.append("</c>\n");
                }
            }
            
            xml.append("</row>\n");
        }
        
        xml.append("</sheetData>\n");
        xml.append("</worksheet>");
        
        zipOut.write(xml.toString().getBytes("UTF-8"));
        zipOut.closeEntry();
    }
    
    private String getColumnName(int colIndex) {
        StringBuilder columnName = new StringBuilder();
        int num = colIndex + 1;
        
        while (num > 0) {
            int remainder = (num - 1) % 26;
            columnName.insert(0, (char) ('A' + remainder));
            num = (num - 1) / 26;
        }
        
        return columnName.toString();
    }
    
    private String escapeXml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace("\"", "&quot;")
                   .replace("'", "&apos;");
    }
    
    private double dateToExcelSerial(String dateStr) {
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            Date date = sdf.parse(dateStr);
            
            // Excel epoch is December 30, 1899
            Calendar excelEpoch = Calendar.getInstance();
            excelEpoch.set(1899, Calendar.DECEMBER, 30, 0, 0, 0);
            excelEpoch.set(Calendar.MILLISECOND, 0);
            
            long diff = date.getTime() - excelEpoch.getTimeInMillis();
            return diff / (1000.0 * 60 * 60 * 24);
        } catch (Exception e) {
            return 0;
        }
    }
    
    private double dateTimeToExcelSerial(String dateTimeStr) {
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Date date = sdf.parse(dateTimeStr);
            
            // Excel epoch is December 30, 1899
            Calendar excelEpoch = Calendar.getInstance();
            excelEpoch.set(1899, Calendar.DECEMBER, 30, 0, 0, 0);
            excelEpoch.set(Calendar.MILLISECOND, 0);
            
            long diff = date.getTime() - excelEpoch.getTimeInMillis();
            return diff / (1000.0 * 60 * 60 * 24);
        } catch (Exception e) {
            return 0;
        }
    }
    
    private double timeToExcelSerial(String timeStr) {
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
            Date time = sdf.parse(timeStr);
            
            Calendar cal = Calendar.getInstance();
            cal.setTime(time);
            
            int hours = cal.get(Calendar.HOUR_OF_DAY);
            int minutes = cal.get(Calendar.MINUTE);
            int seconds = cal.get(Calendar.SECOND);
            
            // Convert to fraction of a day
            return (hours * 3600 + minutes * 60 + seconds) / 86400.0;
        } catch (Exception e) {
            return 0;
        }
    }
}
