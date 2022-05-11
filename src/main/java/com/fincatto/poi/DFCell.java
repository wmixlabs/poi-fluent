package com.fincatto.poi;

import org.apache.poi.ss.usermodel.Cell;

public class DFCell {
    private final Cell cell;

    public DFCell(Cell cell) {
        this.cell = cell;
    }

    public DFCell withValue(final String value) {
        cell.setCellValue(value);
        return this;
    }
}
