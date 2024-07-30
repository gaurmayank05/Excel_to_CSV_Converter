package org.developer;

public class Main {
    public static void main(String[] args) {
        Excel2CSV csvConverter = new Excel2CSV();
        String configurableExcel = "CSD_TO_CSV.xlsx";
        String inputExcel = "CSD_Internal.xlsx";
        try {
            csvConverter.excel2CSV(configurableExcel, inputExcel);
            System.out.println("EXCEL CONVERTED INTO CSV SUCCESSFULLY.");
        } catch (Exception e) {
            //noinspection CallToPrintStackTrace
            e.printStackTrace();
        }
    }
}