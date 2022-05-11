package com.fincatto.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class DFRow {

    private final Row row;

    public DFRow(final Row row) {
        this.row = row;
    }

    public DFCell withCell(final int pos) {
        Cell cel = row.getCell(pos);
        if (cel == null) {
            cel = row.createCell(pos);
        }
        return new DFCell(cel);
    }
}
