package br.com.wmixvideo.poi;

import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

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

    public boolean hasGroupedRows() {
        final Map<String, Integer> groupCouters = new HashMap<>();
        for (WMXRow row : getRows()) {
            if (row.getGroup() != null && !row.getGroup().isBlank()) {
                int groupCount = (groupCouters.getOrDefault(row.getGroup(), 0)) + 1;
                if (groupCount > 1) {
                    return true;
                }
                groupCouters.put(row.getGroup(), groupCount);
            }
        }
        return false;
    }
}
