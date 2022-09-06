package br.com.wmixvideo.poi;

import java.util.ArrayList;
import java.util.List;

public class WMXRow {

    private final List<WMXCell> cells;
    private final WMXSheet parent;
    private String group;
    private String subGroup;

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

    public WMXRow withSubGroup(String subGroup) {
        this.subGroup = subGroup;
        return this;
    }

    public String getGroup() {
        return group;
    }

    public String getSubGroup() {
        return subGroup;
    }

    public WMXSheet and() {
        return this.parent;
    }
}
