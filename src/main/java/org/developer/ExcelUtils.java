package org.developer;

import org.apache.poi.ss.usermodel.*;
import java.io.InputStream;

public class ExcelUtils {
    /**
     * Retrieves an InputStream for a resource file located in the classpath.
     *
     * @param resourceName the name of the resource file within the classpath
     * @return an InputStream for the specified resource file, or null if the resource is not found
     */
    protected InputStream getResourceAsStream(String resourceName) {
        return (getClass().getClassLoader().getResourceAsStream(resourceName));
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
    protected Sheet getSheet (Workbook workbook, ConfigurableExcel parameters){
        Sheet sheet;
        if (workbook.getNumberOfSheets() == 1) {
            sheet = workbook.getSheetAt(0);
        }else {
            sheet = workbook.getSheet(parameters.getSheetName());
        }
        return sheet;
    }

    /**
     * Calculates the maximum column index in the given sheet within the specified sheet of an Excel.
     *
     * @param sheet     is used to store a particular sheet
     * @return The maximum column index found in the specified row range of the sheet, minus one.
     */

    protected int getMaxColumn(Sheet sheet) {
        int maxCol = 0;
        for (Row row : sheet) {
            int lastCol = (row != null) ? row.getLastCellNum() : 0;
            maxCol = Math.max(maxCol, lastCol);
        }
        return maxCol;
    }

    /**
     * Retrieves the string value from the specified Excel cell.
     * Handles different cell types (STRING, NUMERIC, BOOLEAN, FORMULA) and formats.
     *
     * @param cell The Excel Cell object from which to retrieve the value.
     * @return A string representation of the cell value.
     */
    protected String getCellValueasString(Cell cell) {
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
}