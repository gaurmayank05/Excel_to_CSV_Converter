package org.developer;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class Excel2PDF {

    public static void main(String[] args) {
        String excelFilePath = "D://sourceFolder//CSD - Internal - Copy.xlsx";
        String pdfFilePath = "D://convertedPDF//CSD.pdf";
        try {
            convertExcelToPDF(excelFilePath, pdfFilePath);
            System.out.println("Excel file converted to PDF successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void convertExcelToPDF(String excelFilePath, String pdfFilePath) throws IOException, DocumentException {
        try (FileInputStream openExcelFile = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(openExcelFile);
             FileOutputStream pdfFile = new FileOutputStream(pdfFilePath)) {

            Document document = new Document();
            PdfWriter.getInstance(document, pdfFile);
            document.open();

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                PdfPTable table = new PdfPTable(maxColumn(sheet));

                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                            Cell cell = row.getCell(cellIndex);
                            if (cell != null) {
                                String cellValue = getCellValueAsString(cell);
                                PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue, getCellStyle(cell)));
                                setCellAlignment(cell, pdfCell);
                                setBackgroundColor(cell, pdfCell);
                                table.addCell(pdfCell);
                            } else {
                                table.addCell(new PdfPCell());
                            }
                        }
                    }
                }
                applyMergedRegions(sheet, table);
                document.add(table);
                if (sheetIndex < workbook.getNumberOfSheets() - 1) {
                    document.newPage();
                }
            }
            document.close();
        }
    }

    public static int maxColumn(Sheet sheet) {
        int maxCol = 0;
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                int colNum = row.getLastCellNum();
                if (colNum > maxCol) maxCol = colNum;
            }
        }
        return maxCol;
    }

    public static Font getCellStyle(Cell cell) {
        Font font = new Font();
        CellStyle cellStyle = cell.getCellStyle();
        org.apache.poi.ss.usermodel.Font cellFont = cell.getSheet().getWorkbook().getFontAt(cellStyle.getFontIndexAsInt());

        if (cellFont.getItalic()) {
            font.setStyle(Font.ITALIC);
        }

        if (cellFont.getStrikeout()) {
            font.setStyle(Font.STRIKETHRU);
        }

        if (cellFont.getUnderline() == 1) {
            font.setStyle(Font.UNDERLINE);
        }

        font.setSize(cellFont.getFontHeightInPoints());

        if (cellFont.getBold()) {
            font.setStyle(Font.BOLD);
        }

        String fontName = cellFont.getFontName();
        if (FontFactory.isRegistered(fontName)) {
            font.setFamily(fontName);
        } else {
            font.setFamily("Helvetica");
        }
        return font;
    }

    public static void setCellAlignment(Cell cell, PdfPCell cellPdf) {
        CellStyle cellStyle = cell.getCellStyle();
        switch (cellStyle.getAlignment()) {
            case LEFT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_LEFT);
                break;
            case CENTER:
                cellPdf.setHorizontalAlignment(Element.ALIGN_CENTER);
                break;
            case JUSTIFY:
            case FILL:
                cellPdf.setHorizontalAlignment(Element.ALIGN_JUSTIFIED);
                break;
            case RIGHT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_RIGHT);
                break;
        }

        switch (cellStyle.getVerticalAlignment()) {
            case TOP:
                cellPdf.setVerticalAlignment(Element.ALIGN_TOP);
                break;
            case CENTER:
                cellPdf.setVerticalAlignment(Element.ALIGN_MIDDLE);
                break;
            case BOTTOM:
                cellPdf.setVerticalAlignment(Element.ALIGN_BOTTOM);
                break;
            default:
                cellPdf.setVerticalAlignment(Element.ALIGN_UNDEFINED);
        }
        cellPdf.setRotation(cellStyle.getRotation());
    }

    public static void setBackgroundColor(Cell cell, PdfPCell cellPdf) {
        short bgColorIndex = cell.getCellStyle().getFillForegroundColor();
        if (bgColorIndex != IndexedColors.AUTOMATIC.getIndex()) {
            XSSFColor bgColor = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
            if (bgColor != null) {
                byte[] rgb = bgColor.getRGB();
                if (rgb != null && rgb.length == 3) {
                    cellPdf.setBackgroundColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                }
            }
        }
    }

    private static String getCellValueAsString(Cell cell) {
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    private static void applyMergedRegions(Sheet sheet, PdfPTable table) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress region : mergedRegions) {
            int startRow = region.getFirstRow();
            int endRow = region.getLastRow();
            int startCol = region.getFirstColumn();
            int endCol = region.getLastColumn();
            try{
                if (startRow < table.getRows().size() && startCol < table.getNumberOfColumns()) {
                    int rowSpan = endRow - startRow + 1;
                    int colSpan = endCol - startCol + 1;

                    PdfPCell mergedCell = table.getRow(startRow).getCells()[startCol];
                    if (mergedCell != null) {
                        mergedCell.setRowspan(rowSpan);
                        mergedCell.setColspan(colSpan);

                        for (int i = 0; i < rowSpan; i++) {
                            for (int j = 0; j < colSpan; j++) {
                                if (i != 0 || j != 0) {
                                    PdfPCell cellToClear = table.getRow(startRow + i).getCells()[startCol + j];
                                    if (cellToClear != null) {
                                        cellToClear.setBorder(Rectangle.NO_BORDER);
                                    }
                                }
                            }
                        }
                    }
                }
            }catch (Exception e){
                System.out.println(sheet.getSheetName()+" Merged region out of table bounds: " + region.formatAsString());
            }
        }
    }
}