package org.developer;

public class ConfigurableExcel {
    private int startRow;
    private int endRow;
    private final int startColumn;
    private int endColumn;
    private final String sheetName;
    private final String sheetPath;
    private final boolean isTranspose;
    private final boolean isComment;
    private final String sheetRange;
    private final boolean isDeleteAvailable;

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public int getEndColumn() {
        return endColumn;
    }

    public void setEndColumn(int endColumn) {
        this.endColumn = endColumn;
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getSheetPath() {
        return sheetPath;
    }

    public boolean isTranspose() {
        return isTranspose;
    }

    public boolean isComment() {
        return isComment;
    }

    public String getSheetRange() {
        return sheetRange;
    }

    public boolean isDeleteAvailable() {
        return isDeleteAvailable;
    }

    public ConfigurableExcel(int startRow, int endRow, int startColumn, int endColumn, String sheetName, String sheetPath, boolean isTranspose, boolean isComment, String sheetRange, boolean isDeleteAvailable) {
        this.startRow = startRow;
        this.endRow = endRow;
        this.startColumn = startColumn;
        this.endColumn = endColumn;
        this.sheetName = sheetName;
        this.sheetPath = sheetPath;
        this.isTranspose = isTranspose;
        this.isComment = isComment;
        this.sheetRange = sheetRange;
        this.isDeleteAvailable = isDeleteAvailable;
    }
}