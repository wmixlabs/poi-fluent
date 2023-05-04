package br.com.wmixvideo.poi;

import java.util.ArrayList;
import java.util.List;

public class WMXSheet {

    private final String name;
    private final WMXSpreadsheet parent;
    private final List<WMXRow> rows;
    private int freezeCols, freezeRows;
    private boolean autoSizeColumns;
    private WMXCellRange autoFilterRange;
    private boolean collapseRowGroups;

    public WMXSheet(final String name, final WMXSpreadsheet parent) {
        this.name = name;
        this.parent = parent;
        this.rows = new ArrayList<>();
    }

    public WMXRow withRow() {
        final WMXRow row = new WMXRow(this);
        this.rows.add(row);
        return row;
    }

    public WMXSheet freeze(final int cols, final int rows) {
        this.freezeCols = cols;
        this.freezeRows = rows;
        return this;
    }

    public WMXSheet unfreeze() {
        this.freezeCols = 0;
        this.freezeRows = 0;
        return this;
    }

    public String getName() {
        return name;
    }

    public List<WMXRow> getRows() {
        return rows;
    }

    public int getFreezeCols() {
        return freezeCols;
    }

    public int getFreezeRows() {
        return freezeRows;
    }

    public WMXSheet withAutoSizeColumns(final boolean autoSizeColumns) {
        this.autoSizeColumns = autoSizeColumns;
        return this;
    }

    public WMXSheet withAutoFilter(final int firstRow, final int firstColumn, final int lastRow, final int lastColumn) {
        this.autoFilterRange = new WMXCellRange(firstRow, firstColumn, lastRow, lastColumn);
        return this;
    }

    public boolean isAutoSizeColumns() {
        return autoSizeColumns;
    }

    public WMXCellRange getAutoFilterRange() {
        return autoFilterRange;
    }

    public WMXSpreadsheet and() {
        return parent;
    }


    public boolean isCollapseRowGroups() {
        return collapseRowGroups;
    }

    public WMXSheet withCollapseRowGroups(boolean collapseRowGroups) {
        this.collapseRowGroups = collapseRowGroups;
        return this;
    }
}
