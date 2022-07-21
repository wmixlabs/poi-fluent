package com.fincatto.poi;

import org.apache.poi.ss.usermodel.IndexedColors;

public class DFCell<T> {
    private T value;
    private DFStyle style;
    private String formula;
    private String comment;
    private String link;
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

    public DFCell<T> withFontColor(final IndexedColors color) {
        this.style.setFontColor(color);
        return this;
    }

    public DFCell<T> withFontSize(final Short size) {
        this.style.setFontSize(size);
        return this;
    }

    public DFCell<T> withBackgroundColor(final IndexedColors color) {
        this.style.setBackgroundColor(color);
        return this;
    }

    public DFCell<T> withComment(final String comment) {
        this.comment = comment;
        return this;
    }

    public String getComment() {
        return comment;
    }

    public DFCell<T> withMergedCells(final int size) {
        this.mergedCells = size;
        return this;
    }

    public int getMergedCells() {
        return mergedCells;
    }

    public DFCell<T> withMergedRows(final int size) {
        this.mergedRows = size;
        return this;
    }

    public int getMergedRows() {
        return mergedRows;
    }

    public DFCell<T> withDataFormat(final String dataFormat) {
        this.getStyle().setDataFormat(dataFormat);
        return this;
    }

    public DFCell<T> withDataFormat(final short dataForatBuiltin) {
        this.getStyle().setDataFormatBuiltin(dataForatBuiltin);
        return this;
    }

    public DFCell<T> withLink(final String link) {
        this.link = link;
        return this;
    }

    public String getLink() {
        return link;
    }

    public DFCell<T> withFormula(final String formula) {
        this.formula = formula;
        return this;
    }

    public String getFormula() {
        return this.formula;
    }
}
