package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelToCSV {
    /**
     * Converts specified sheets and ranges from an Excel file to CSV format based on a configurable Excel file.
     * @param configurableExcel the name to the configurable Excel file containing the conversion parameters
     * @param inputExcel the name to the input Excel file to be converted
     * @throws Exception if an error occurs during the conversion process
     */
    public void ExcelToCSVConverter(String configurableExcel, String inputExcel) throws Exception {
        ZipDirectory zipDirectory = new ZipDirectory();
        String tempFolder = zipDirectory.createTempDirectory("tempCSV");
        String zipDestinationFolder = "D://CSV.zip";
        ConfigurableExcel excelQueryParameters = new ConfigurableExcel(0, -1, 1, -1, null, null, false, true, null, false);
        InputStream configurableExcelPath = getResourceAsStream(configurableExcel);
        List<List<String>> excelConfigurationList = queryExcelData(configurableExcelPath, excelQueryParameters);
        List<ConfigurableExcel> queryConfigList = fillSheetParameter(excelConfigurationList);
        validateSheetAndPath(queryConfigList, excelConfigurationList,inputExcel);
        for (ConfigurableExcel parameters : queryConfigList) {
            List<List<String>> excelData;
            InputStream inputExcelPath = getResourceAsStream(inputExcel);
            if (parameters.getSheetRange().isEmpty() || parameters.getSheetRange() == null) {
                excelData = queryExcelData(inputExcelPath, parameters);
            } else {
                excelData = specificRange(inputExcel, parameters);
            }
            if (parameters.isDeleteAvailable()) {
                excelData.add(addDeleteColumn(excelData));
            }
            if (parameters.isTranspose()) {
                excelData = transposeData(excelData);
            }
            String csvFilePath = createDirectory(tempFolder, parameters);
            writeCSV(parameters, excelData, csvFilePath);
            inputExcelPath.close();
        }
        zipDirectory.zipFolder(tempFolder, zipDestinationFolder);
        zipDirectory.deleteTempDirectory(tempFolder);
        configurableExcelPath.close();
    }
    /**
     * Validates the sheet names and paths in the configuration.
     * Throws exceptions if any inconsistencies are found.
     *
     * @param queryConfigList    The list of configurable Excel parameters.
     * @param configExcelList    The configuration data from the Excel file.
     * @throws Exception If any validation fails.
     */
    private void validateSheetAndPath(List<ConfigurableExcel> queryConfigList, List<List<String>> configExcelList, String inputExcel) throws Exception {
        for (List<String> rowData : configExcelList) {
            if (rowData.stream().allMatch(cellData -> cellData.trim().isEmpty())) {
                throw new Exception("CONFIGURABLE EXCEL SHEET CONTAINS BLANK ROWS");
            }
        }
        for (ConfigurableExcel parameters : queryConfigList) {
            boolean isSheetNameEmpty = parameters.getSheetName() == null || parameters.getSheetName().trim().isEmpty();
            boolean isSheetPathEmpty = parameters.getSheetPath() == null || parameters.getSheetPath().trim().isEmpty();
            if (isSheetNameEmpty && !isSheetPathEmpty) {
                throw new Exception("CSD SHEET DOES NOT EXIST BUT CSV DIRECTORY PATH EXISTS: " + parameters.getSheetPath());
            }
            if (!isSheetNameEmpty && isSheetPathEmpty) {
                throw new Exception(parameters.getSheetName() + " CSD SHEET EXISTS BUT CSV DIRECTORY PATH DOES NOT EXIST");
            }
            if (isSheetNameEmpty) {
                throw new Exception("CSD SHEET AND CSV DIRECTORY PATH DOES NOT EXIST");
            }
        }
        Map<String, List<Integer>> multipleBlankRowsMap = new HashMap<>();
        Map<String, Integer> singleBlankRowMap = new HashMap<>();
        List<String> whitespaceErrors = new ArrayList<>();
        List<String> blankRowWhitespaceErrors = new ArrayList<>();
        try (InputStream inputExcelPath = getResourceAsStream(inputExcel); Workbook workbook = new XSSFWorkbook(inputExcelPath)) {
            for (ConfigurableExcel parameters : queryConfigList) {
                String sheetName = parameters.getSheetName();
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    throw new Exception("SHEET DOES NOT EXIST: " + sheetName);
                }
                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        boolean isRowBlank = true;
                        boolean hasWhitespaceInBlankCell = false;
                        for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                            Cell cell = row.getCell(cellIndex);
                            if (cell != null) {
                                String cellValue = cell.toString();
                                if (cellValue.trim().isEmpty() && !cellValue.isEmpty()) {
                                    hasWhitespaceInBlankCell = true;
                                    whitespaceErrors.add("Whitespace in Sheet: " + sheetName + " at Row: " + (rowIndex + 1) + " Column: " + (cellIndex + 1));
                                }
                                if (cell.getCellType() != CellType.BLANK && !cellValue.trim().isEmpty()) {
                                    isRowBlank = false;
                                }
                            }
                        }
                        if (isRowBlank) {
                            if (multipleBlankRowsMap.containsKey(sheetName)) {
                                multipleBlankRowsMap.get(sheetName).add(rowIndex + 1);
                            } else if (singleBlankRowMap.containsKey(sheetName)) {
                                List<Integer> rowList = new ArrayList<>();
                                rowList.add(singleBlankRowMap.remove(sheetName));
                                rowList.add(rowIndex + 1);
                                multipleBlankRowsMap.put(sheetName, rowList);
                            } else {
                                singleBlankRowMap.put(sheetName, rowIndex + 1);
                            }
                            if (hasWhitespaceInBlankCell) {
                                blankRowWhitespaceErrors.add("Whitespace in blank row in Sheet: " + sheetName + " at Row: " + (rowIndex + 1));
                            }
                        }
                    }
                }
            }
        }
        StringBuilder sb = new StringBuilder();
        if (!multipleBlankRowsMap.isEmpty()) {
            sb.append("Multiple Blank Rows: ");
            for (Map.Entry<String, List<Integer>> entry : multipleBlankRowsMap.entrySet()) {
                sb.append("\nSheet: ").append(entry.getKey()).append(" Rows: ").append(entry.getValue());
            }
            sb.append("\n");
        }
        if (!singleBlankRowMap.isEmpty()) {
            if (!sb.isEmpty()) {
                sb.append("\n");
            }
            sb.append("Single Blank Row: ");
            for (Map.Entry<String, Integer> entry : singleBlankRowMap.entrySet()) {
                sb.append("\nSheet: ").append(entry.getKey()).append(" Row: ").append(entry.getValue());
            }
            sb.append("\n");
        }
        if (!whitespaceErrors.isEmpty()) {
            if (!sb.isEmpty()) {
                sb.append("\n");
            }
            sb.append("Whitespace errors found in existing rows: ");
            for (String error : whitespaceErrors) {
                sb.append("\n").append(error);
            }
            sb.append("\n");
        }
        if (!blankRowWhitespaceErrors.isEmpty()) {
            if (!sb.isEmpty()) {
                sb.append("\n");
            }
            sb.append("Whitespace errors found in blank rows: ");
            for (String error : blankRowWhitespaceErrors) {
                sb.append("\n").append(error);
            }
        }
        if (!sb.isEmpty()) {
            throw new Exception(sb.toString());
        }
    }

    /**
     * Queries the data from an Excel file based on the provided parameters.
     *
     * @param getExcelPath the InputStream of the Excel files either configurable Excel file or data Excel file
     * @param parameters the configurable Excel parameters for querying the data
     * @return a list of lists, where each inner list represents a row of data from the Excel file
     */
    private List<List<String>> queryExcelData(InputStream getExcelPath, ConfigurableExcel parameters){
        List<List<String>> excelData = new ArrayList<>();
        try (getExcelPath; Workbook workbook = new XSSFWorkbook(getExcelPath)) {
            Sheet sheet = getSheet(workbook, parameters);
            if (parameters.isTranspose() && (parameters.getSheetRange().isEmpty() || parameters.getSheetRange() == null)) {
                parameters.setStartRow(2);
            }
            if (parameters.getEndRow() == -1) {
                parameters.setEndRow(sheet.getLastRowNum());
            }
            if (parameters.getEndColumn() == -1) {
                parameters.setEndColumn(maxColumn(sheet, parameters));
            }
            if (parameters.isComment()) {
                parameters.setEndColumn(parameters.getEndColumn() + 1);
            }
            for (int rowIndex = parameters.getStartRow(); rowIndex <= parameters.getEndRow(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                List<String> rowData = new ArrayList<>();
                for (int cellIndex = parameters.getStartColumn(); cellIndex < parameters.getEndColumn(); cellIndex++) {
                    if (row != null) {
                        Cell cell = row.getCell(cellIndex);
                        if (cell != null) {
                            rowData.add(getCellValueasString(cell).trim());
                        }
                    }
                }
                excelData.add(rowData);
            }
        } catch (IOException e) {
            //noinspection CallToPrintStackTrace
            e.printStackTrace();
        }
        return excelData;
    }

    /**
     * Retrieves the appropriate sheet from the workbook based on the provided parameters.
     *
     * @param workbook the workbook containing the sheets based on needed
     * @param parameters the configurable Excel parameters containing the sheet name
     * @return            A Sheet object based on the conditions:
     *                    - If the workbook has exactly one sheet, returns that sheet (only for Configurable Excel).
     *                    - If the workbook has more than one sheet, returns the sheet with the name (for CSD Excel).
     */
    private Sheet getSheet (Workbook workbook, ConfigurableExcel parameters){
        Sheet sheet;
        if (workbook.getNumberOfSheets() == 1) {
            sheet = workbook.getSheetAt(0);
        }else {
            sheet = workbook.getSheet(parameters.getSheetName());
        }
        return sheet;
    }

    /**
     * Transposes the given two-dimensional list of strings (Excel data).
     *
     * @param excelData The original two-dimensional list of strings to transpose.
     * @return A transposed two-dimensional list of strings, where rows become columns and vice versa.
     * Returns an empty list if excelData is empty or null.
     */

    private List<List<String>> transposeData(List<List<String>> excelData) {
        List<List<String>> transposedData = new ArrayList<>();
        if (excelData == null || excelData.isEmpty()) {
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
        if (excelData == null || excelData.isEmpty()) return addDeleteColumn;
        int colCount = 0;
        for (List<String> row : excelData) {
            if (row.size() > colCount) colCount = row.size();
        }
        for (int j = 0; j < colCount; j++) {
            if (j == 0) addDeleteColumn.add("deleted");
            else addDeleteColumn.add("False");
        }
        return addDeleteColumn;
    }

    /**
     * Calculates the maximum column index in the given sheet within the specified sheet of an Excel.
     *
     * @param sheet     is used to store a particular sheet
     * @param parameter The configurableExcel object containing parameters for particular sheet.
     * @return The maximum column index found in the specified row range of the sheet, minus one.
     */

    private int maxColumn(Sheet sheet, ConfigurableExcel parameter) {
        int endColumn = 0;
        for (int rowIndex = parameter.getStartRow(); rowIndex <= parameter.getEndRow(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getLastCellNum() > endColumn) endColumn = row.getLastCellNum();
        }
        return endColumn - 1;
    }

    /**
     * Writes the provided Excel data to a CSV file based on the parameters.
     *
     * @param parameters The configurableExcel object containing parameters for particular sheet.
     * @param excelData the data to be written to the CSV file
     * @param csvFilePath is used to store the path of temporary folder
     * @throws IOException if an error occurs while writing the CSV file
     */
    private void writeCSV(ConfigurableExcel parameters, List<List<String>> excelData, String csvFilePath) throws IOException {

        if (parameters.getSheetPath() != null) {
            try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFilePath), StandardCharsets.UTF_8))) {
                // Write the standardized header to the CSV file
                standardizedHeader(writer, excelData);
                for (int rowIndex = 1; rowIndex < excelData.size(); rowIndex++) {
                    List<String> row = excelData.get(rowIndex);
                    for (int i = 0; i < row.size(); i++) {
                        if (row.get(i) != null) writer.write(especialCharacters(row.get(i)));
                        else writer.append("");
                        if (i < row.size() - 1) writer.append(",");
                    }
                    if (rowIndex != excelData.size()-1) writer.newLine();
                }
            }
        }
    }

    /**
     * Creates a directory based on the provided configurableExcel parameter's sheet path.
     *
     * @param destinationFolder is used to store the path of destinationFolder
     *                          ( for below method it's used to store the path of temporary folder)
     * @param parameter The configurableExcel object containing sheet information.
     * @return The absolute path of the directory where the file will be saved,
     *         or an empty string if the sheet path is null.
     */
    private String createDirectory(String destinationFolder, ConfigurableExcel parameter) {
        String absolutePath = "";
        if (parameter.getSheetPath() != null) {
            File file = new File(parameter.getSheetPath());
            String directoryPath = file.getParent();
            String csvFile = file.getName();
            File folder = new File(destinationFolder + File.separator + directoryPath);
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
     * Standardizes and writes the header row of Excel data to the specified BufferedWriter.
     *
     * @param writer    The BufferedWriter to write the standardized header data.
     * @param excelData The two-dimensional list representing Excel data, where the first list is assumed to be the header row.
     * @throws IOException If an I/O error occurs while writing to the BufferedWriter.
     */
    private void standardizedHeader(BufferedWriter writer, List<List<String>> excelData) throws IOException {
        List<String> excelHeaderData = excelData.get(0);
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
     * Each range is processed separately and concatenated into a single two-dimensional list of strings.
     * for example: (3, 10-15), 8 etc.
     *
     * @param inputExcel the path to the input Excel file
     * @param parameters the configurable Excel parameters containing the sheet range and other settings
     * @return a list of lists, where each inner list represents a row of data from the specified ranges in the Excel file
     * @throws IOException if an error occurs while reading the Excel file
     */
    private List<List<String>> specificRange(String inputExcel, ConfigurableExcel parameters) throws IOException {
        int startRow, endRow;
        List<List<String>> excelData = null;
        // Split the sheet range parameter into individual ranges
        String[] range = parameters.getSheetRange().split(",");
        for (String rangeIndex : range) {
            // Load the input Excel file from the resource folder
            InputStream inputExcelPath = getResourceAsStream(inputExcel);
            // Parse the start row from the range and adjust for zero-based indexing
            startRow = Integer.parseInt(rangeIndex.split("-")[0].trim()) - 1;
            // Parse the end row from the range if it exists, otherwise set it to the start row
            if (rangeIndex.contains("-")) {
                endRow = Integer.parseInt(rangeIndex.split("-")[1].trim()) - 1;
            }
            else endRow=startRow;

            parameters.setStartRow(startRow);
            parameters.setEndRow(endRow);
            List<List<String>> tempExcelData = queryExcelData(inputExcelPath, parameters);
            if (excelData == null) excelData = tempExcelData;
            else excelData.addAll(tempExcelData);
            inputExcelPath.close();
        }
        return excelData;
    }

    /**
     * Retrieves the string value from the specified Excel cell.
     * Handles different cell types (STRING, NUMERIC, BOOLEAN, FORMULA) and formats.
     *
     * @param cell The Excel Cell object from which to retrieve the value.
     * @return A string representation of the cell value.
     */
    private String getCellValueasString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue().toString();
                else return String.valueOf((int) cell.getNumericCellValue());
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
     * Each inner list represents data for a configurableExcel object.
     *
     * @param configurableExcelData The two-dimensional list of strings containing data for configurableExcel objects.
     * @return A list of configurableExcel objects populated with data from configurableExcelData.
     */
    private List<ConfigurableExcel> fillSheetParameter(List<List<String>> configurableExcelData) {
        List<ConfigurableExcel> queryConfigList = new ArrayList<>();
        for (int rowIndex = 1; rowIndex < configurableExcelData.size(); rowIndex++) {
            List<String> rowData = configurableExcelData.get(rowIndex);
            ConfigurableExcel parameters = new ConfigurableExcel(0, -1, 1, -1, rowData.get(0),
                    rowData.get(1), Boolean.parseBoolean(rowData.get(2)), Boolean.parseBoolean(rowData.get(3)), rowData.get(4),
                    Boolean.parseBoolean(rowData.get(5)));
            queryConfigList.add(parameters);
        }
        return queryConfigList;
    }

    /**
     * Retrieves an InputStream for a resource file located in the classpath.
     *
     * @param resourcePath the path to the resource file within the classpath
     * @return an InputStream for the specified resource file, or null if the resource is not found
     */
    private InputStream getResourceAsStream(String resourcePath) {
        return (getClass().getClassLoader().getResourceAsStream(resourcePath));
    }

    public static void main(String[] args) {
        ExcelToCSV csvConverter = new ExcelToCSV();
        String configurableExcel = "CSD_TO_CSV.xlsx";
        String inputExcel = "CSD_Internal.xlsx";
        try {
            csvConverter.ExcelToCSVConverter(configurableExcel, inputExcel);
            System.out.println("EXCEL CONVERTED INTO CSV SUCCESSFULLY.");
        } catch (Exception e) {
            //noinspection CallToPrintStackTrace
            e.printStackTrace();
        }
    }
}