package com.fincatto.poi;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;

public class DFRow {

    private List<DFCell> cells;

    public DFRow() {
        this.cells = new ArrayList<>();
    }

    public <FIELDVALUE> DFCell<FIELDVALUE> withCell(FIELDVALUE value) {
        final DFCell cell = new DFCell(value);
        cells.add(cell);
        return cell;
    }

//    public DFCell withCell(final int pos) {
//        Cell cel = row.getCell(pos);
//        if (cel == null) {
//            cel = row.createCell(pos);
//        }
//        return new DFCell(cel);
//    }


    public List<DFCell> getCells() {
        return cells;
    }
}
