package br.com.wmixvideo.poi;

import java.io.Serializable;

public class WMXCellRange implements Serializable {
    private int firstRow, firstColumn, lastRow, lastColumn;

    public WMXCellRange() {
    }

    public WMXCellRange(int firstRow, int firstColumn, int lastRow, int lastColumn) {
        this.firstRow = firstRow;
        this.firstColumn = firstColumn;
        this.lastRow = lastRow;
        this.lastColumn = lastColumn;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public WMXCellRange setFirstRow(int firstRow) {
        this.firstRow = firstRow;
        return this;
    }

    public int getFirstColumn() {
        return firstColumn;
    }

    public WMXCellRange setFirstColumn(int firstColumn) {
        this.firstColumn = firstColumn;
        return this;
    }

    public int getLastRow() {
        return lastRow;
    }

    public WMXCellRange setLastRow(int lastRow) {
        this.lastRow = lastRow;
        return this;
    }

    public int getLastColumn() {
        return lastColumn;
    }

    public WMXCellRange setLastColumn(int lastColumn) {
        this.lastColumn = lastColumn;
        return this;
    }
}
