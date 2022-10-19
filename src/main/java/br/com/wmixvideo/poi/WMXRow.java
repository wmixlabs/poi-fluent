package br.com.wmixvideo.poi;

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

    public WMXCell withEmptyCell(){
        return this.withCell(null);
    }

    public WMXRow withEmptyCells(final int size){
        for(int i = 0; i< size ; i++){
            this.withCell(null);
        }
        return this;
    }

    public String getGroup() {
        return group;
    }

    public WMXSheet and(){
        return this.parent;
    }

    public int getIndex() {
        int index = 0;
        for (int i = 0; i < this.parent.getRows().indexOf(this); i++) {
            index++;
        }
        return index+1;
    }
}
