package com.fincatto.poi;

import java.util.ArrayList;
import java.util.List;

public class DFRow {

    private List<DFCell> cells;
    private Object agrupador;

    public DFRow() {
        this.cells = new ArrayList<>();
    }

    public <FIELDVALUE> DFCell<FIELDVALUE> withCell(FIELDVALUE value) {
        final DFCell cell = new DFCell(value);
        cells.add(cell);
        return cell;
    }

    public List<DFCell> getCells() {
        return cells;
    }

    public DFRow withAgrupador(Object agrupador) {
        this.agrupador = agrupador;
        return this;
    }

    public Object getAgrupador() {
        return agrupador;
    }
}
