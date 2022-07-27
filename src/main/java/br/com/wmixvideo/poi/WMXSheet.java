package br.com.wmixvideo.poi;

import java.util.ArrayList;
import java.util.List;

class WMXSheet {

    private final String name;
    private final List<WMXRow> rows;
    private int freezeCols, freezeRows;
    private boolean autoSizeColumns;

    public WMXSheet(final String name) {
        this.name = name;
        this.rows = new ArrayList<>();
    }

    public WMXRow withRow() {
        final WMXRow row = new WMXRow();
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

    public boolean isAutoSizeColumns() {
        return autoSizeColumns;
    }
}
