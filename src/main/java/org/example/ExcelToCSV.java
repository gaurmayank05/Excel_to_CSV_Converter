// Added Name
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
     * @param configurableExcelPath used to store the path of Configurable Excel
     * @param inputExcelPath used to store the path of Excel which need to converted into CSV files.
     * @throws IOException If an I/O error occurs while converting the Excel file.
     */

    public void ExcelToCSVConverter(String configurableExcelPath, String inputExcelPath) throws IOException {
        configurableExcel excelQueryParameters = new configurableExcel(0,-1,1,-1,null,null,false,true,"na");
        List<List<String>> excelConfigurationList = queryExcelData(configurableExcelPath, excelQueryParameters);
        List<configurableExcel> queryConfigList = fillSheetParameter(excelConfigurationList);

        for( configurableExcel parameters : queryConfigList  ){
                List<List<String>> excelData;
                if (parameters.getSheetRange().equals("na")){
                    excelData = queryExcelData(inputExcelPath, parameters);
                }else{
                    excelData = specificRange(inputExcelPath, parameters);
                }

                if (parameters.isTranspose()) {
                    excelData = transposeData(excelData);
                }
                writeCSV( parameters, excelData);
        }
    }

    /**
     * Creates a directory based on the provided configurableExcel parameter's sheet path.
     *
     * @param parameter The configurableExcel object containing sheet information.
     * @return The absolute path of the directory where the file will be saved,
     *         or an empty string if the sheet path is null.
     */

    private String createDirectory(configurableExcel parameter) {
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
     * Queries data from an Excel file based on specified parameters.
     *
     * @param getExcelPath The file path of the Excel workbook to query.
     * @param parameters   The configurableExcel object containing parameters for querying data.
     * @return A list of lists representing the queried data from the Excel sheet.
     * @throws IOException If an I/O error occurs while reading the Excel file.
     */

    private List<List<String>> queryExcelData(String getExcelPath, configurableExcel parameters) throws IOException {
        List<List<String>> excelData = new ArrayList<>();
        try (FileInputStream excelFile = new FileInputStream(getExcelPath);
             Workbook workbook = new XSSFWorkbook(excelFile)) {
            Sheet sheet;
            if (workbook.getNumberOfSheets()==1){
                sheet = workbook.getSheetAt(0);
            }else {
                sheet = workbook.getSheet(parameters.getSheetName());
            }
            if (parameters.isTranspose() && parameters.getSheetRange().equals("na")){
                parameters.setStartRow(2);
            }
            if (parameters.getEndRow()==-1){
                parameters.setEndRow(sheet.getLastRowNum());
            }
            if (parameters.getEndColumn()==-1){
                parameters.setEndColumn(maxColumn(sheet,parameters));
            }
            if (parameters.isComment()) {
                parameters.setEndColumn(parameters.getEndColumn()+1);
            }
            for (int rowIndex = parameters.getStartRow(); rowIndex <= parameters.getEndRow(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                List<String> rowData = new ArrayList<>();
                for (int cellIndex = parameters.getStartColumn(); cellIndex < parameters.getEndColumn(); cellIndex++) {
                    if (row != null) {
                        Cell cell = row.getCell(cellIndex);
                        if (cell != null) {
                            rowData.add(getCellValueasString(cell).trim());
                        }else{
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

    public static List<List<String>> transposeData(List<List<String>> excelData ) {

        if(excelData == null)
            return null;

        List<List<String>> transposedData = new ArrayList<>();
        //noinspection ConstantValue
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
     * Calculates the maximum column index in the given sheet within the specified sheet of an Excel.
     *
     * @param sheet     is used to store a particular sheet
     * @param parameter The configurableExcel object containing parameters for particular sheet.
     * @return The maximum column index found in the specified row range of the sheet, minus one.
     */

    private int maxColumn(Sheet sheet, configurableExcel parameter) {
        int endColumn = 0;
        for (int rowIndex = parameter.getStartRow(); rowIndex <= parameter.getEndRow(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getLastCellNum() > endColumn) {
                endColumn = row.getLastCellNum();
            }
        }
        return endColumn-1;
    }

    /**
     *
     * @param parameters The configurableExcel object containing parameters for particular sheet.
     * @param excelData The original two-dimensional list of strings which used to store the Excel Data.
     * @throws IOException If an I/O error occurs while writing the Excel file into CSV file.
     */

    private void writeCSV(configurableExcel parameters, List<List<String>> excelData) throws IOException {

        if (parameters.getSheetPath() != null) {
            String csvFilePath = createDirectory(parameters);

            try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFilePath), StandardCharsets.UTF_8))) {
                standardizedHeader(writer, excelData);
                for (int rowIndex = 1; rowIndex < excelData.size(); rowIndex++) {
                    List<String> row = excelData.get(rowIndex);
                    for (int i = 0; i < row.size(); i++) {
                        if (row.get(i) != null){
                            writer.write(especialCharacters(row.get(i)));
                        } else {
                            writer.append("");
                        }
                        if (i < row.size() - 1) {
                            writer.append(",");
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
     * @param excelData The two-dimensional list representing Excel data, where the first list is assumed to be the header row.
     * @throws IOException If an I/O error occurs while writing to the BufferedWriter.
     */

    private void standardizedHeader(BufferedWriter writer, List<List<String>> excelData) throws IOException {
        List<String> excelHeaderData = excelData.get(0);

        for (int columnIndex = 0; columnIndex < excelHeaderData.size(); columnIndex++) {
            String headerData = excelHeaderData.get(columnIndex);
            if (headerData != null) {
                headerData = headerData.toLowerCase().replaceAll(" ", "_");
            }
            writer.append(headerData != null ? headerData : "");
            if (columnIndex < excelHeaderData.size() - 1) writer.append(",");
        }
        writer.newLine();
    }

    /**
     * Retrieves Excel data from specific row ranges defined in parameters.getSheetRange().
     * Each range is processed separately and concatenated into a single two-dimensional list of strings.
     *
     * @param inputExcelPath The file path of the Excel workbook to query.
     * @param parameters     The configurableExcel object containing parameters for row ranges.
     * @return A two-dimensional list of strings representing the concatenated Excel data from specified row ranges.
     * @throws IOException If an I/O error occurs while reading the Excel file or querying data.
     */

    private List<List<String>> specificRange(String inputExcelPath, configurableExcel parameters) throws IOException {
        int startRow, endRow;
        List<List<String>> excelData = null;
        String[] range = parameters.getSheetRange().split(",");
        for (String rangeIndex : range) {
            startRow = Integer.parseInt(rangeIndex.split("-")[0].trim()) - 1;
            if (rangeIndex.contains("-")){
                endRow = Integer.parseInt(rangeIndex.split("-")[1].trim()) - 1;
            }else{
                endRow = startRow;
            }
            parameters.setStartRow(startRow);
            parameters.setEndRow(endRow);
            List<List<String>> tempExcelData = queryExcelData(inputExcelPath, parameters);
            if (excelData == null) {
                excelData = tempExcelData;
            } else {
                excelData.addAll(tempExcelData);
            }
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

    public static String getCellValueasString(Cell cell) {
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

    public static String especialCharacters(String cellValue) {
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

    public List<configurableExcel> fillSheetParameter(List<List<String>>configurableExcelData) {
        List<configurableExcel> queryConfigList = new ArrayList<>();
        for (int rowIndex = 1; rowIndex < configurableExcelData.size(); rowIndex++) {
            List<String> rowData = configurableExcelData.get(rowIndex);
            configurableExcel parameters = new configurableExcel(0, -1, 1, -1, rowData.get(0), rowData.get(1), Boolean.parseBoolean(rowData.get(2)), Boolean.parseBoolean(rowData.get(3)), rowData.get(4));
            queryConfigList.add(parameters);
        }
        return queryConfigList;
    }

    public static void main(String[] args) {
        ExcelToCSV csvConverter = new ExcelToCSV();
        String configurableExcelPath = "D://sourceFolder/CSD_TO_CSV.xlsx";
        String inputExcelPath = "D://sourceFolder//CSD - Internal.xlsx";
        try {
            csvConverter.ExcelToCSVConverter(configurableExcelPath, inputExcelPath);
            System.out.println("Excel is converted into CSV.");
        } catch (IOException e) {
            throw new NullPointerException();
        }
    }
}