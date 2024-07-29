package org.developer;

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


public class Excel2CSV {
    /**
     * Converts specified sheets and ranges from an Excel file to CSV format based on a configurable Excel file.
     * @param configurableExcel the name to the configurable Excel file containing the conversion parameters
     * @param inputExcel the name to the input Excel file to be converted
     * @throws Exception if an error occurs during the conversion process
     */
    public void excel2CSV(String configurableExcel, String inputExcel) throws Exception {
        ExcelUtils excelUtils = new ExcelUtils();
        ZipDirectory zipDirectory = new ZipDirectory();
        String tempFolder = zipDirectory.createTempDirectory("tempCSV");
        String zipDestinationFolder = "D://CSV.zip";
        ConfigurableExcel excelQueryParameters = new ConfigurableExcel(0, -1, 1, -1, null, null, false, true, null, false);
        InputStream configurableExcelPath = excelUtils.getResourceAsStream(configurableExcel);
        List<List<String>> excelConfigurationList = queryExcelData(configurableExcelPath, excelQueryParameters);
        List<ConfigurableExcel> queryConfigList = fillSheetParameter(excelConfigurationList);
        validateSheetAndPath(queryConfigList, excelConfigurationList,inputExcel);
        for (ConfigurableExcel parameters : queryConfigList) {
            List<List<String>> excelData;
            InputStream inputExcelPath = excelUtils.getResourceAsStream(inputExcel);
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
        ExcelUtils excelUtils = new ExcelUtils();
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
        // * Maps to keep track of blank rows and errors
        Map<String, List<Integer>> multipleBlankRowsMap = new HashMap<>();
        Map<String, Integer> singleBlankRowMap = new HashMap<>();
        List<String> blankRowWhitespaceErrors = new ArrayList<>();
        try (InputStream inputExcelPath = excelUtils.getResourceAsStream(inputExcel);
             Workbook workbook = new XSSFWorkbook(inputExcelPath)) {
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
                        StringBuilder rowWhitespaceErrors = new StringBuilder();
                        for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                            Cell cell = row.getCell(cellIndex);
                            if (cell != null) {
                                String cellValue = cell.toString();
                                if (cellValue.trim().isEmpty() && !cellValue.isEmpty()) {
                                    hasWhitespaceInBlankCell = true;
                                    rowWhitespaceErrors.append(" and Column: ").append(cellIndex + 1).append(" ");
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
                                blankRowWhitespaceErrors.add("Whitespace in New blank Row in Sheet: " + sheetName + " at Row: " + (rowIndex + 1) + rowWhitespaceErrors);
                            }
                        }
                    }
                }
            }
        }
        // * Build error messages
        StringBuilder errorSummary = new StringBuilder();

        if (!multipleBlankRowsMap.isEmpty()) {
            errorSummary.append("Multiple Blank Rows: ");
            for (Map.Entry<String, List<Integer>> entry : multipleBlankRowsMap.entrySet()) {
                errorSummary.append("\nSheet: ").append(entry.getKey()).append(" Rows: ").append(entry.getValue());
            }
            errorSummary.append("\n");
        }
        if (!singleBlankRowMap.isEmpty()) {
            if (!errorSummary.isEmpty()) {
                errorSummary.append("\n");
            }
            errorSummary.append("Single Blank Row: ");
            for (Map.Entry<String, Integer> entry : singleBlankRowMap.entrySet()) {
                errorSummary.append("\nSheet: ").append(entry.getKey()).append(" Row:").append(entry.getValue());
            }
            errorSummary.append("\n");
        }
        if (!blankRowWhitespaceErrors.isEmpty()) {
            if (!errorSummary.isEmpty()) {
                errorSummary.append("\n");
            }
            errorSummary.append("Whitespace errors found in blank rows: ");
            for (String error : blankRowWhitespaceErrors) {
                errorSummary.append("\n").append(error);
            }
        }
        if (!errorSummary.isEmpty()) {
            throw new Exception(errorSummary.toString());
        }
    }

    /**
     * Queries the data from an Excel file based on the provided parameters.
     *
     * @param getExcelPath the InputStream of the Excel files either configurable Excel file or data Excel file
     * @param parameters the configurable Excel parameters for querying the data
     * @return a list of lists, where each inner list represents a row of data from the Excel file
     */
    public List<List<String>> queryExcelData(InputStream getExcelPath, ConfigurableExcel parameters){
        ExcelUtils excelUtils = new ExcelUtils();
        List<List<String>> excelData = new ArrayList<>();
        try (getExcelPath; Workbook workbook = new XSSFWorkbook(getExcelPath)) {
            Sheet sheet = excelUtils.getSheet(workbook, parameters);
            if (parameters.isTranspose() && (parameters.getSheetRange().isEmpty() || parameters.getSheetRange() == null)) {
                parameters.setStartRow(2);
            }
            if (parameters.getEndRow() == -1) {
                parameters.setEndRow(sheet.getLastRowNum());
            }
            if (parameters.getEndColumn() == -1) {
                parameters.setEndColumn(excelUtils.maxColumn(sheet, parameters));
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
                            rowData.add(excelUtils.getCellValueasString(cell).trim());
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
    protected String createDirectory(String destinationFolder, ConfigurableExcel parameter) {
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
        ExcelUtils excelUtils = new ExcelUtils();
        int startRow, endRow;
        List<List<String>> excelData = null;
        // Split the sheet range parameter into individual ranges
        String[] range = parameters.getSheetRange().split(",");
        for (String rangeIndex : range) {
            // Load the input Excel file from the resource folder
            InputStream inputExcelPath = excelUtils.getResourceAsStream(inputExcel);
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
}