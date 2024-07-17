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
     *
     * @throws IOException If an I/O error occurs while converting the Excel file.
     */

    public void ExcelToCSVConverter(String configurableExcel, String inputExcel) throws Exception {
        ConfigurableExcel excelQueryParameters = new ConfigurableExcel(0, -1, 1, -1, null, null, false, true, null, false);

        InputStream configurableExcelPath = getResourceAsStream(configurableExcel);

        List<List<String>> excelConfigurationList = queryExcelData(configurableExcelPath, excelQueryParameters);
        List<ConfigurableExcel> queryConfigList = fillSheetParameter(excelConfigurationList);
        for (ConfigurableExcel parameters : queryConfigList) {
            validateSheetAndPath(parameters);
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
            writeCSV(parameters, excelData);
        }
    }

    private void validateSheetAndPath(ConfigurableExcel parameters) throws Exception {
        boolean isSheetNameEmpty = parameters.getSheetName() == null || parameters.getSheetName().isEmpty();
        boolean isSheetPathEmpty = parameters.getSheetPath() == null || parameters.getSheetPath().isEmpty();
        if (isSheetNameEmpty && !isSheetPathEmpty) {
            throw new Exception("CSD SHEET DOES NOT EXIST BUT CSV DIRECTORY PATH EXISTS:" + parameters.getSheetPath());
        }
        if (!isSheetNameEmpty && isSheetPathEmpty) {
            throw new Exception(parameters.getSheetName()+" CSD SHEET  EXIST BUT CSV DIRECTORY PATH  NOT EXISTS "  );
        }
        if (isSheetNameEmpty) {
            throw new Exception("CSD SHEET AND CSV DIRECTORY PATH DOES NOT EXIST");
        }
    }

    /**
     * Queries data from an Excel file based on specified parameters.
     *
     * @param getExcelPath The file path of the Excel workbook to query.
     * @param parameters   The configurableExcel object containing parameters for querying data.
     * @return A list of lists representing the queried data from the Excel sheet.
     * @throws IOException If an I/O error occurs while reading the Excel file.
     */

    private List<List<String>> queryExcelData(InputStream getExcelPath, ConfigurableExcel parameters) throws IOException {
        List<List<String>> excelData = new ArrayList<>();
        try {
            Workbook workbook = new XSSFWorkbook(getExcelPath);
            Sheet sheet = getSheet(workbook, parameters);
            if(parameters.isTranspose() && (parameters.getSheetRange().isEmpty() || parameters.getSheetRange()==null))
            {
                parameters.setStartRow(2);
            }
            if (parameters.getEndRow()==-1){
                parameters.setEndRow(sheet.getLastRowNum());
            }
            if (parameters.getEndColumn()==-1){
                parameters.setEndColumn(maxColumn(sheet,parameters));
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
        }catch (IOException e){
            e.printStackTrace();
        }
        return excelData;
    }

    /**
     * Retrieves a Sheet from the provided Workbook based on the specified parameters.
     *
     * @param workbook    The Workbook from which to retrieve the Sheet.
     * @param parameters  The ConfigurableExcel parameters containing additional configuration.
     * @return            A Sheet object based on the conditions:
     *                      - If the workbook has exactly one sheet, returns that sheet (only for Configurable Excel).
     *                      - If the workbook has more than one sheet, returns the sheet with the name (for CSD Excel).
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
     * Write the data of an Excel sheet into CSV format.
     *
     * @param parameters The configurableExcel object containing parameters for particular sheet.
     * @param excelData The original two-dimensional list of strings which used to store the Excel Data.
     * @throws IOException If an I/O error occurs while writing the Excel file into CSV file.
     */

    private void writeCSV(ConfigurableExcel parameters, List<List<String>> excelData) throws IOException {

        if (parameters.getSheetPath() != null) {
            String csvFilePath = createDirectory(parameters);
            try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFilePath), StandardCharsets.UTF_8))) {
                standardizedHeader(writer, excelData);
                for (int rowIndex = 1; rowIndex < excelData.size(); rowIndex++) {
                    List<String> row = excelData.get(rowIndex);
                    for (int i = 0; i < row.size(); i++) {
                        if (row.get(i) != null) writer.write(especialCharacters(row.get(i)));
                        else writer.append("");
                        if (i < row.size() - 1) writer.append(",");
                    }
                    writer.newLine();
                }
            }
        }
    }

    /**
     * Creates a directory based on the provided configurableExcel parameter's sheet path.
     *
     * @param parameter The configurableExcel object containing sheet information.
     * @return The absolute path of the directory where the file will be saved,
     *         or an empty string if the sheet path is null.
     */
    private String createDirectory(ConfigurableExcel parameter) {
        String outputDirectory = "D://";
        String absolutePath = "";
        if (parameter.getSheetPath() != null) {
            File file = new File(parameter.getSheetPath());
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
     * @param parameters     The configurableExcel object containing parameters for row ranges.
     * @return A two-dimensional list of strings representing the concatenated Excel data from specified row ranges.
     * @throws IOException If an I/O error occurs while reading the Excel file or querying data.
     */

    private List<List<String>> specificRange(String inputExcel, ConfigurableExcel parameters) throws IOException {
        int startRow, endRow;
        List<List<String>> excelData = null;
        String[] range = parameters.getSheetRange().split(",");
        for (String rangeIndex : range) {
            InputStream inputExcelPath = getResourceAsStream(inputExcel);
            startRow = Integer.parseInt(rangeIndex.split("-")[0].trim()) - 1;
            if (rangeIndex.contains("-")) {
                endRow = Integer.parseInt(rangeIndex.split("-")[1].trim()) - 1;
            }
            else endRow=startRow;
            parameters.setStartRow(startRow);
            parameters.setEndRow(endRow);
            List<List<String>> tempExcelData = queryExcelData(inputExcelPath, parameters);
            if (excelData == null) excelData = tempExcelData;
            else excelData.addAll(tempExcelData);
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