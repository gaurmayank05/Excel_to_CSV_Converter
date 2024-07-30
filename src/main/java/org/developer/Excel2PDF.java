package org.developer;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import com.itextpdf.text.pdf.BaseFont;

public class Excel2PDF {

    public static void main(String[] args) {
        ExcelUtils excelUtils = new ExcelUtils();
        String excelFile = "CSD_Internal.xlsx";
        String pdfFilePath = "D://convertedPDF//CSD.pdf";
        try {
            convertExcelToPDF(excelUtils.getResourceAsStream(excelFile), pdfFilePath);
            System.out.println("Excel file converted to PDF successfully.");
        } catch (Exception e) {
            //noinspection CallToPrintStackTrace
            e.printStackTrace();
        }
    }

    public static void convertExcelToPDF(InputStream excelFilePath, String pdfFilePath) throws IOException, DocumentException {
        try (excelFilePath; Workbook workbook = new XSSFWorkbook(excelFilePath);
             FileOutputStream pdfFile = new FileOutputStream(pdfFilePath)) {
            ExcelUtils excelUtils = new ExcelUtils();
            Document document = new Document(PageSize.A4.rotate());
            PdfWriter.getInstance(document, pdfFile);
            document.open();
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                int maxColumns = excelUtils.getMaxColumn(sheet);
                float pdfWidth = document.getPageSize().getWidth() - document.leftMargin() - document.rightMargin();
                PdfPTable table = createTable(sheet, maxColumns, pdfWidth);
                table.completeRow();
                document.add(table);
                if (sheetIndex < workbook.getNumberOfSheets() - 1) {
                    document.newPage();
                }
            }
            document.close();
        }
    }

    private static PdfPTable createTable(Sheet sheet, int maxColumns, float pdfWidth) throws DocumentException {
        PdfPTable table = new PdfPTable(maxColumns);
        table.setWidthPercentage(100);
        table.setWidths(getScaledColumnWidths(sheet, maxColumns, pdfWidth));
        PdfPCell pdfCell;
        for (Row row : sheet) {
            if (row != null) {
                for (int cellIndex = 0; cellIndex < maxColumns; cellIndex++){
                    Cell cell = row.getCell(cellIndex);
                    if (cell != null){
                        pdfCell = createPdfCell(cell, maxColumns);
                    }else{
                        pdfCell = new PdfPCell();
                    }
                    table.addCell(pdfCell);
                }
            }
        }
        applyMergedRegions(sheet, table);
        return table;
    }

    private static float[] getScaledColumnWidths(Sheet sheet, int maxColumns, float pdfWidth) {
        float[] columnWidths = new float[maxColumns];
        float totalWidth = 0;
        for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++) {
            columnWidths[columnIndex] = sheet.getColumnWidthInPixels(columnIndex);
            totalWidth += columnWidths[columnIndex];
        }
        float scale = pdfWidth / totalWidth;
        for (int columnIndex = 0; columnIndex < columnWidths.length; columnIndex++) {
            columnWidths[columnIndex] *= scale;
        }
        return columnWidths;
    }

    private static PdfPCell createPdfCell(Cell cell, int maxColumns) {
        ExcelUtils excelUtils = new ExcelUtils();
        String cellValue = excelUtils.getCellValueasString(cell);
        Font font = getFont(cell, maxColumns);
        PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue, font));
        setCellAlignment(cell, pdfCell);
        setBackgroundColor(cell, pdfCell);
        return pdfCell;
    }

    private static Font getFont(Cell cell, int maxColumns) {
        CellStyle cellStyle = cell.getCellStyle();
        //noinspection deprecation
        org.apache.poi.ss.usermodel.Font cellFont = cell.getSheet().getWorkbook().getFontAt(cellStyle.getFontIndexAsInt());
        Font font = new Font();
        font.setColor(getFontColor(cellFont));
        font.setSize(applyFontSize(maxColumns));
        font.setStyle(getFontStyle(cellFont));
        font.setFamily(getFontFamily(cellFont));
            try {
                BaseFont baseFont = BaseFont.createFont("arial-unicode-ms.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                font = new Font(baseFont, font.getSize(), font.getStyle(), font.getColor());
            } catch (DocumentException | IOException e) {
                e.printStackTrace();
            }
        return font;
    }
    private static BaseColor getFontColor(org.apache.poi.ss.usermodel.Font cellFont) {
        if (cellFont instanceof XSSFFont) {
            XSSFColor xssfColor = ((XSSFFont) cellFont).getXSSFColor();
            if (xssfColor != null) {
                byte[] rgb = xssfColor.getRGB();
                if (rgb != null){
                    return new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
                }
            }
        }
        return BaseColor.BLACK;
    }

    private static int applyFontSize(int maxColumns) {
        if (maxColumns > 5) return 6;
        return 11;
    }

    private static String getFontFamily(org.apache.poi.ss.usermodel.Font cellFont) {
        String fontName = cellFont.getFontName();
        return FontFactory.isRegistered(fontName) ? fontName : FontFactory.HELVETICA;
    }

    private static int getFontStyle(org.apache.poi.ss.usermodel.Font cellFont) {
        int style = Font.NORMAL;
        if (cellFont.getBold()) {
            style = Font.BOLD;
        }
        if (cellFont.getItalic()) {
            style = Font.ITALIC;
        }
        if (cellFont.getStrikeout()) {
            style = Font.STRIKETHRU;
        }
        if (cellFont.getUnderline() == 1) {
            style = Font.UNDERLINE;
        }
        return style;
    }

    private static void setCellAlignment(Cell cell, PdfPCell pdfCell) {
        CellStyle cellStyle = cell.getCellStyle();
        switch (cellStyle.getAlignment()) {
            case LEFT:
                pdfCell.setHorizontalAlignment(Element.ALIGN_LEFT);
                break;
            case CENTER:
                pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                break;
            case RIGHT:
                pdfCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
                break;
            default:
                pdfCell.setHorizontalAlignment(Element.ALIGN_UNDEFINED);
                break;
        }

        switch (cellStyle.getVerticalAlignment()) {
            case TOP:
                pdfCell.setVerticalAlignment(Element.ALIGN_TOP);
                break;
            case CENTER:
                pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                break;
            case BOTTOM:
                pdfCell.setVerticalAlignment(Element.ALIGN_BOTTOM);
                break;
            default:
                pdfCell.setVerticalAlignment(Element.ALIGN_UNDEFINED);
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

                    for (int rowIndex = 0; rowIndex <= endRow - startRow; rowIndex++) {
                        for (int cellIndex = 0; cellIndex <= endCol - startCol; cellIndex++) {
                            if (rowIndex != 0 || cellIndex != 0) {
                                PdfPCell cellToClear = table.getRow(startRow + rowIndex).getCells()[startCol + cellIndex];
                                if (cellToClear != null) {
                                    cellToClear.setBorder(Rectangle.NO_BORDER);
                                }
                            }
                        }
                    }
                }
            } catch (Exception e) {
                System.out.println("Error in merging regions: " + region.formatAsString() + " in sheet: " + sheet.getSheetName());
                e.printStackTrace();
            }
        }
    }
}