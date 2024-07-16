package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelToCSV {
    /**
     * Converts Excel files to CSV based on configurable parameters.
     *
     * @param configurableExcelPath The input stream for the configurable Excel file.
     * @param inputExcelPath        The input stream for the Excel file to convert.
     * @throws Exception If any error occurs during the conversion process.
     */
    public void ExcelToCSVConverter(InputStream configurableExcelPath, InputStream inputExcelPath) throws Exception {
        List<ConfigurableExcel> queryConfigList = readConfigurableExcel(configurableExcelPath);

        for (ConfigurableExcel parameters : queryConfigList) {
            List<List<String>> excelData;

            if (parameters.getSheetRange() == null || parameters.getSheetRange().isEmpty()) {
                excelData = queryExcelData(inputExcelPath, parameters);
            } else {
                excelData = specificRange(inputExcelPath, parameters);
            }

            if (parameters.isDeleteAvailable()) {
                excelData.add(addDeleteColumn(excelData));
            }

            if (parameters.isTranspose()) {
                excelData = transposeData(excelData);
            }

            writeCSV(parameters, excelData);
        }
    }

    /**
     * Reads configurable parameters from the Excel file.
     *
     * @param configurableExcelPath The input stream for the configurable Excel file.
     * @return A list of ConfigurableExcel objects parsed from the Excel file.
     * @throws IOException If any I/O error occurs during reading the Excel file.
     */
    private List<ConfigurableExcel> readConfigurableExcel(InputStream configurableExcelPath) throws IOException {
        List<List<String>> excelConfigurationList = queryExcelData(configurableExcelPath, new ConfigurableExcel(0, -1, 1, -1, null, null,
                false, true, null, false));
        return fillSheetParameter(excelConfigurationList);
    }

    /**
     * Creates a directory for the CSV file based on the sheet path.
     *
     * @param parameters The configurableExcel object containing parameters for directory creation.
     * @return The absolute path of the directory where the file will be saved.
     */
    private String createDirectory(ConfigurableExcel parameters) {
        String outputDirectory = "D://";
        String absolutePath = "";

        if (parameters.getSheetPath() != null) {
            File file = new File(parameters.getSheetPath());
            String directoryPath = file.getParent();
            String csvFile = file.getName();
            File folder = new File(outputDirectory + File.separator + directoryPath);

            if (!folder.exists()) {
                //noinspection ResultOfMethodCallIgnored
                folder.mkdirs();
            }
            absolutePath = folder.getAbsolutePath() + File.separator + csvFile;
        } else {
            System.out.println("Warning: Sheet path is null. Returning empty directory path.");
        }

        return absolutePath;
    }


    /**
     * Queries data from an Excel file based on specified parameters.
     *
     * @param excelPath   The input stream for the Excel file.
     * @param parameters  The configurableExcel object containing parameters for querying data.
     * @return A list of lists representing the queried data from the Excel sheet.
     * @throws IOException If any I/O error occurs during reading the Excel file.
     */
    private List<List<String>> queryExcelData(InputStream excelPath, ConfigurableExcel parameters) throws IOException {
        List<List<String>> excelData = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(excelPath)) {
            Sheet sheet = getSheet(workbook, parameters);
            int endColumn = getMaxColumn(sheet, parameters);

            for (int rowIndex = parameters.getStartRow(); rowIndex <= parameters.getEndRow(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                List<String> rowData = new ArrayList<>();
                for (int cellIndex = parameters.getStartColumn(); cellIndex <= endColumn; cellIndex++) {
                    if (row != null) {
                        Cell cell = row.getCell(cellIndex);
                        if (cell != null) {
                            rowData.add(getCellValueAsString(cell).trim());
                        } else {
                            rowData.add("");
                        }
                    }
                }
                excelData.add(rowData);
            }
        }
        return excelData;
    }

    /**
     * Transposes the given two-dimensional list of strings (Excel data).
     *
     * @param excelData The original two-dimensional list of strings to transpose.
     * @return A transposed two-dimensional list of strings, where rows become columns and vice versa.
     *         Returns an empty list if excelData is empty or null.
     */

    private List<List<String>> transposeData(List<List<String>> excelData ) {

        List<List<String>> transposedData = new ArrayList<>();
        if (excelData == null || excelData.isEmpty()){
            return transposedData;
        }
        int rowCount = excelData.size();
        int colCount = 0;
        for (List<String> row : excelData) {
            if (row.size() > colCount) {
                colCount = row.size();
            }
        }
        for (int j = 0; j < colCount; j++) {
            List<String> row = new ArrayList<>();
            for (int i = 0; i < rowCount; i++) {

                if (j < excelData.get(i).size()) {
                    row.add(excelData.get(i).get(j));
                }
            }
            transposedData.add(row);
        }
        return transposedData;
    }

    /**
     * Add an extra Deleted Column in some file if needed
     * @param excelData The original two-dimensional list of strings in which an extra Delete column need to be added.
     * @return A list of Excel Data with an extra Delete Column along its default value 'False'.
     */
    private List<String> addDeleteColumn(List<List<String>> excelData) {
        List<String> addDeleteColumn = new ArrayList<>();
        if (excelData == null || excelData.isEmpty()){
            return addDeleteColumn;
        }
        int colCount = 0;
        for (List<String> row : excelData) {
            if (row.size() > colCount) {
                colCount = row.size();
            }
        }
        for (int j = 0; j < colCount; j++) {
            if (j==0) {
                addDeleteColumn.add("deleted");
            }else{
                addDeleteColumn.add("False");
            }
        }
        return addDeleteColumn;
    }

    /**
     * Retrieves the maximum column index in the given sheet within the specified parameters.
     *
     * @param sheet      The sheet object to query.
     * @param parameters The configurableExcel object containing parameters for the sheet.
     * @return The maximum column index found in the specified row range of the sheet.
     */
    private int getMaxColumn(Sheet sheet, ConfigurableExcel parameters) {
        int endColumn = 0;
        for (int rowIndex = parameters.getStartRow(); rowIndex <= parameters.getEndRow(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getLastCellNum() > endColumn) {
                endColumn = row.getLastCellNum();
            }
        }
        return endColumn - 1;
    }

    /**
     * Retrieves the sheet from the workbook based on the specified parameters.
     *
     * @param workbook   The workbook object containing the sheets.
     * @param parameters The configurableExcel object containing parameters for the sheet.
     * @return The sheet object to query data from.
     */
    private Sheet getSheet(Workbook workbook, ConfigurableExcel parameters) {
        if (workbook.getNumberOfSheets() == 1) {
            return workbook.getSheetAt(0);
        } else {
            return workbook.getSheet(parameters.getSheetName());
        }
    }

    /**
     * Converts Excel data to CSV format and writes it to a file.
     *
     * @param parameters The configurableExcel object containing parameters for CSV writing.
     * @param excelData  The two-dimensional list of strings representing Excel data.
     * @throws IOException If any I/O error occurs during writing the CSV file.
     */
    private void writeCSV(ConfigurableExcel parameters, List<List<String>> excelData) throws IOException {
        if (parameters.getSheetPath() != null) {
            String csvFilePath = createDirectory(parameters);
            try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFilePath), StandardCharsets.UTF_8))) {
                standardizedHeader(writer, excelData.get(0));
                for (int i = 1; i < excelData.size(); i++) {
                    List<String> row = excelData.get(i);
                    for (int j = 0; j < row.size(); j++) {
                        writer.write(especialCharacters(row.get(j)));
                        if (j < row.size() - 1) {
                            writer.write(",");
                        }
                    }
                    writer.newLine();
                }
            }
        }
    }

    /**
     * Standardizes and writes the header row of Excel data to the specified BufferedWriter.
     *
     * @param writer    The BufferedWriter to write the standardized header data.
     * @param excelHeaderData is a list representing Excel data of first row.
     * @throws IOException If an I/O error occurs while writing to the BufferedWriter.
     */

    private void standardizedHeader(BufferedWriter writer, List<String> excelHeaderData) throws IOException {
        for (int columnIndex = 0; columnIndex < excelHeaderData.size(); columnIndex++) {
            String headerData = excelHeaderData.get(columnIndex);
            if (headerData != null) {

                headerData = headerData.replace("*", "").toLowerCase().replaceAll("\\s+", "_")
                        .replaceAll("_+$", "");
            }
            writer.append(headerData != null ? headerData : "");
            if (columnIndex < excelHeaderData.size() - 1) writer.append(",");
        }
        writer.newLine();
    }

    /**
     * Retrieves Excel data from specific row ranges defined in parameters.getSheetRange().
     *
     * @param inputExcelPath The input stream for the Excel file.
     * @param parameters     The configurableExcel object containing parameters for row ranges.
     * @return A two-dimensional list of strings representing the concatenated Excel data from specified row ranges.
     * @throws IOException If any I/O error occurs during reading the Excel file.
     */
    private List<List<String>> specificRange(InputStream inputExcelPath, ConfigurableExcel parameters) throws IOException {
        List<List<String>> excelData = new ArrayList<>();
        if (parameters.getSheetRange().contains(",")) {
            String[] ranges = parameters.getSheetRange().split(",");
            for (String range : ranges) {
                if (range.contains("-")) {
                    String[] startEnd = range.split("-");
                    int startRow = Integer.parseInt(startEnd[0].trim()) - 1;
                    int endRow = startEnd.length > 1 ? Integer.parseInt(startEnd[1].trim()) - 1 : startRow;
                    parameters.setStartRow(startRow);
                    parameters.setEndRow(endRow);
                    excelData.addAll(queryExcelData(inputExcelPath, parameters));
                }
            }
        }
        return excelData;
    }

    /**
     * Retrieves the string value from the specified Excel cell.
     *
     * @param cell The Excel Cell object from which to retrieve the value.
     * @return A string representation of the cell value.
     */
    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    /**
     * Escapes special characters in the given cell value for CSV format.
     * Special characters include double quotes, commas, newline characters, single quotes, slashes, and backslashes.
     *
     * @param cellValue The original cell value to escape special characters from.
     * @return The escaped cell value formatted for CSV.
     */

    private String especialCharacters(String cellValue) {
        cellValue = cellValue.replaceAll("\"", "\"\"");
        String regex = "[,\\n'/\\\\\"]";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(cellValue);
        if (matcher.find()) {
            cellValue = "\"" + cellValue + "\"";
        }
        return cellValue;
    }

    /**
     * Converts a two-dimensional list of strings into a list of configurableExcel objects.
     *
     * @param configurableExcelData The two-dimensional list of strings containing data for configurableExcel objects.
     * @return A list of configurableExcel objects populated with data from configurableExcelData.
     */
    private List<ConfigurableExcel> fillSheetParameter(List<List<String>> configurableExcelData) {
        List<ConfigurableExcel> queryConfigList = new ArrayList<>();
        for (int i = 1; i < configurableExcelData.size(); i++) {
            List<String> rowData = configurableExcelData.get(i);
            ConfigurableExcel parameters = new ConfigurableExcel(0, -1, 1, -1, rowData.get(0), rowData.get(1),
                    Boolean.parseBoolean(rowData.get(2)), Boolean.parseBoolean(rowData.get(3)), rowData.get(4),
                    Boolean.parseBoolean(rowData.get(5)));
            queryConfigList.add(parameters);
        }
        return queryConfigList;
    }

    private InputStream getResourceAsStream(String resourcePath) {
        return (getClass().getClassLoader().getResourceAsStream(resourcePath));
    }

    public static void main(String[] args) {
        ExcelToCSV csvConverter = new ExcelToCSV();
        InputStream configurableExcelPath = csvConverter.getResourceAsStream("CSD_TO_CSV.xlsx");
        InputStream inputExcelPath = csvConverter.getResourceAsStream("CSD_Internal.xlsx");
        try {
            csvConverter.ExcelToCSVConverter(configurableExcelPath, inputExcelPath);
            System.out.println("EXCEL To CSV CONVERSION SUCCESSFULLY.");
        } catch (Exception e) {
            throw new NullPointerException();
        }
    }
}