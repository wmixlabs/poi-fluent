package com.fincatto.poi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

class DFSheet {
    private final Sheet sheet;

    public DFSheet(final Sheet sheet) {
        this.sheet = sheet;
    }

    public DFRow withRow(final int pos) {
        Row row = this.sheet.getRow(pos);
        if (row == null) {
            row = this.sheet.createRow(pos);
        }
        return new DFRow(row);
    }

    public DFSheet freeze(final int rows, final int cols) {
        sheet.createFreezePane(rows, cols);
        return this;
    }

    public DFSheet unfreeze() {
        sheet.createFreezePane(0, 0);
        return this;
    }
}
