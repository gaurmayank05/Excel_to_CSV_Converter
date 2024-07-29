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
        String excelFilePath = "D://sourceFolder//CSD - Internal.xlsx";
        String pdfFilePath = "D://convertedPDF//CSD.pdf";
        try {
            convertExcelToPDF(excelFilePath, pdfFilePath, 10);
            System.out.println("Excel file converted to PDF successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void convertExcelToPDF(String excelFilePath, String pdfFilePath, int defaultFontSize) throws IOException, DocumentException {
        try (FileInputStream excelFile = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(excelFile);
             FileOutputStream pdfFile = new FileOutputStream(pdfFilePath)) {
            Document document = new Document(PageSize.A4.rotate());
            PdfWriter.getInstance(document, pdfFile);
            document.open();
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                int maxColumns = getMaxColumns(sheet);
                PdfPTable table = createTable(sheet, maxColumns, document.getPageSize().getWidth() - document.leftMargin() - document.rightMargin(), defaultFontSize);
                document.add(table);
                if (sheetIndex < workbook.getNumberOfSheets() - 1) {
                    document.newPage();
                }
            }
            document.close();
        }
    }

    private static int getMaxColumns(Sheet sheet) {
        int maxCol = 0;
        for (Row row : sheet) {
            int lastCol = (row != null) ? row.getLastCellNum() : 0;
            maxCol = Math.max(maxCol, lastCol);
        }
        return maxCol;
    }

    private static PdfPTable createTable(Sheet sheet, int maxColumns, float pdfWidth, int defaultFontSize) throws DocumentException {
        PdfPTable table = new PdfPTable(maxColumns);
        table.setWidthPercentage(100);
        table.setWidths(getScaledColumnWidths(sheet, maxColumns, pdfWidth));
        for (Row row : sheet) {
            for (int cellIndex = 0; cellIndex < maxColumns; cellIndex++) {
                Cell cell = (row != null) ? row.getCell(cellIndex) : null;
                PdfPCell pdfCell = (cell != null) ? createPdfCell(cell, defaultFontSize) : new PdfPCell();
                table.addCell(pdfCell);
            }
        }
        applyMergedRegions(sheet, table);
        return table;
    }

    private static float[] getScaledColumnWidths(Sheet sheet, int maxColumns, float pdfWidth) {
        float[] columnWidths = new float[maxColumns];
        float totalWidth = 0;
        for (int i = 0; i < maxColumns; i++) {
            columnWidths[i] = sheet.getColumnWidthInPixels(i);
            totalWidth += columnWidths[i];
        }
        float scale = pdfWidth / totalWidth;
        for (int i = 0; i < columnWidths.length; i++) {
            columnWidths[i] *= scale;
        }
        return columnWidths;
    }

    private static PdfPCell createPdfCell(Cell cell, int defaultFontSize) {
        String cellValue = getCellValueAsString(cell);
        Font font = getFont(cell, defaultFontSize);
        PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue, font));
        setCellAlignment(cell, pdfCell);
        setBackgroundColor(cell, pdfCell);
        return pdfCell;
    }

    private static String getCellValueAsString(Cell cell) {
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    private static Font getFont(Cell cell, int defaultFontSize) {
        CellStyle cellStyle = cell.getCellStyle();
        org.apache.poi.ss.usermodel.Font cellFont = cell.getSheet().getWorkbook().getFontAt(cellStyle.getFontIndexAsInt());
        Font font = new Font();
        font.setSize(6);
        font.setStyle(getFontStyle(cellFont));
        font.setFamily(getFontFamily(cellFont));

        return font;
    }

    private static String getFontFamily(org.apache.poi.ss.usermodel.Font cellFont) {
        String fontName = cellFont.getFontName();
        return FontFactory.isRegistered(fontName) ? fontName : FontFactory.HELVETICA;
    }

    private static int getFontStyle(org.apache.poi.ss.usermodel.Font cellFont) {
        int style = Font.NORMAL;
        if (cellFont.getBold()) style |= Font.BOLD;
        if (cellFont.getItalic()) style |= Font.ITALIC;
        if (cellFont.getStrikeout()) style |= Font.STRIKETHRU;
        if (cellFont.getUnderline() == 1) style |= Font.UNDERLINE;
        return style;
    }

    private static void setCellAlignment(Cell cell, PdfPCell pdfCell) {
        CellStyle cellStyle = cell.getCellStyle();
        switch (cellStyle.getAlignment()) {
            case LEFT -> pdfCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            case CENTER -> pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            case RIGHT -> pdfCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            default -> pdfCell.setHorizontalAlignment(Element.ALIGN_UNDEFINED);
        }

        switch (cellStyle.getVerticalAlignment()) {
            case TOP -> pdfCell.setVerticalAlignment(Element.ALIGN_TOP);
            case CENTER -> pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            case BOTTOM -> pdfCell.setVerticalAlignment(Element.ALIGN_BOTTOM);
            default -> pdfCell.setVerticalAlignment(Element.ALIGN_UNDEFINED);
        }
        pdfCell.setRotation(cellStyle.getRotation());
    }

    private static void setBackgroundColor(Cell cell, PdfPCell pdfCell) {
        short bgColorIndex = cell.getCellStyle().getFillForegroundColor();
        if (bgColorIndex != IndexedColors.AUTOMATIC.getIndex()) {
            XSSFColor bgColor = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
            if (bgColor != null) {
                byte[] rgb = bgColor.getRGB();
                if (rgb != null && rgb.length == 3) {
                    pdfCell.setBackgroundColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                }
            }
        }
    }

    private static void applyMergedRegions(Sheet sheet, PdfPTable table) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress region : mergedRegions) {
            int startRow = region.getFirstRow();
            int endRow = region.getLastRow();
            int startCol = region.getFirstColumn();
            int endCol = region.getLastColumn();

            try {
                PdfPCell mergedCell = table.getRow(startRow).getCells()[startCol];
                if (mergedCell != null) {
                    mergedCell.setRowspan(endRow - startRow + 1);
                    mergedCell.setColspan(endCol - startCol + 1);

                    for (int i = 0; i <= endRow - startRow; i++) {
                        for (int j = 0; j <= endCol - startCol; j++) {
                            if (i != 0 || j != 0) {
                                PdfPCell cellToClear = table.getRow(startRow + i).getCells()[startCol + j];
                                if (cellToClear != null) {
                                    cellToClear.setBorder(Rectangle.NO_BORDER);
                                }
                            }
                        }
                    }
                }
            } catch (Exception e) {
                System.out.println("Error in merging regions: " + region.formatAsString() + " in sheet: " + sheet.getSheetName());
            }
        }
    }
}
