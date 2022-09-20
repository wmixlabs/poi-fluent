package br.com.wmixvideo.poi;

import org.apache.poi.ss.formula.functions.T;

import java.util.ArrayList;
import java.util.List;

public class WMXRow {

    private final List<WMXCell> cells;
    private final WMXSheet parent;
    private String group;
    private boolean hiddenRow;

    public WMXRow(final WMXSheet sheet) {
        parent = sheet;
        this.cells = new ArrayList<>();
    }

    public <FIELDVALUE> WMXCell<FIELDVALUE> withCell(FIELDVALUE value) {
        final WMXCell cell = new WMXCell(value, this);
        cells.add(cell);
        return cell;
    }

    public List<WMXCell> getCells() {
        return cells;
    }

    public WMXRow withGroup(String group) {
        this.group = group;
        return this;
    }
    public WMXRow withHiddenRow(final boolean hiddenRow){
        this.hiddenRow = hiddenRow;
        return this;
    }

    public boolean isHiddenRow() {
        return hiddenRow;
    }

    public String getGroup() {
        return group;
    }

    public WMXSheet and(){
        return this.parent;
    }
}
