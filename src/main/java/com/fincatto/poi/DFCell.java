package com.fincatto.poi;

import org.apache.poi.ss.usermodel.IndexedColors;

public class DFCell<T> {
    private T value;
    private DFStyle style;
    private String comment;
    private int mergedCells;
    private int mergedRows;

    public DFCell(T value) {
        this.value = value;
        this.style = new DFStyle();
        this.mergedCells = 0;
        this.mergedRows = 0;
    }

    public DFCell<T> withValue(final T value) {
        this.value = value;
        return this;
    }

    public T getValue() {
        return value;
    }


    public DFStyle getStyle() {
        return style;
    }

    public DFCell<T> title() {
        this.style.setFontBold(true);
        this.style.setBackgroundColor(IndexedColors.GREY_25_PERCENT);
        return this;
    }

    public DFCell<T> bold() {
        this.style.setFontBold(true);
        return this;
    }

    public DFCell<T> withFontFamily(final String fontFamily) {
        this.style.setFont(fontFamily);
        return this;
    }

    public DFCell<T> withBackgroundColor(IndexedColors color) {
        this.style.setBackgroundColor(color);
        return this;
    }

    public DFCell<T> withComment(String comment){
        this.comment = comment;
        return this;
    }

    public String getComment() {
        return comment;
    }

    public DFCell<T> withMergedCells(final int size){
        this.mergedCells = size;
        return this;
    }

    public int getMergedCells() {
        return mergedCells;
    }

    public DFCell<T> withMergedRows(final int size){
        this.mergedRows = size;
        return this;
    }

    public int getMergedRows() {
        return mergedRows;
    }
}
