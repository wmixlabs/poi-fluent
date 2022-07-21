package com.fincatto.poi;

import java.util.ArrayList;
import java.util.List;

class DFSheet {

    private final String name;
    private List<DFRow> rows;
    private Integer freezeCols, freezeRows;
    private boolean autoSizeColumns;

    public DFSheet(final String name) {
        this.name = name;
        this.rows = new ArrayList<>();
    }

    public DFRow withRow(){
        final DFRow row = new DFRow();
        this.rows.add(row);
        return row;
    }

    public DFSheet freeze(final int cols, final int rows) {
        this.freezeCols = cols;
        this.freezeRows = rows;
        return this;
    }

    public DFSheet unfreeze() {
        this.freezeCols = null;
        this.freezeRows = null;
        return this;
    }

    public String getName() {
        return name;
    }

    public List<DFRow> getRows() {
        return rows;
    }

    public Integer getFreezeCols() {
        return freezeCols;
    }

    public Integer getFreezeRows() {
        return freezeRows;
    }

    public DFSheet withAutoSizeColumns(final boolean autoSizeColumns){
        this.autoSizeColumns = autoSizeColumns;
        return this;
    }

    public boolean isAutoSizeColumns() {
        return autoSizeColumns;
    }
}
