package br.com.wmixvideo.poi;

import org.apache.poi.ss.usermodel.BorderStyle;

import java.util.ArrayList;
import java.util.List;

public class WMXRow {

    private final List<WMXCell> cells;
    private String group;

    public WMXRow() {
        this.cells = new ArrayList<>();
    }

    public <FIELDVALUE> WMXCell<FIELDVALUE> withCell(FIELDVALUE value) {
        final WMXCell cell = new WMXCell(value);
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

    public String getGroup() {
        return group;
    }
}
