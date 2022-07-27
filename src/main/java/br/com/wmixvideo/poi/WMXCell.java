package br.com.wmixvideo.poi;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

public class WMXCell<T> {

    private T value;
    private final WMXStyle style;
    private String formula, comment, link;
    private int mergedColumns, mergedRows;

    public WMXCell(T value) {
        this.value = value;
        this.style = new WMXStyle();
    }

    public WMXCell<T> withValue(final T value) {
        this.value = value;
        return this;
    }

    public T getValue() {
        return value;
    }

    public WMXStyle getStyle() {
        return style;
    }

    public WMXCell<T> title() {
        this.style.setFontBold(true);
        this.style.setBackgroundColor(IndexedColors.GREY_25_PERCENT);
        return this;
    }

    public WMXCell<T> bold() {
        this.style.setFontBold(true);
        return this;
    }

    public WMXCell<T> withFontFamily(final String fontFamily) {
        this.style.setFont(fontFamily);
        return this;
    }

    public WMXCell<T> withFontColor(final IndexedColors color) {
        this.style.setFontColor(color);
        return this;
    }

    public WMXCell<T> withFontSize(final Short size) {
        this.style.setFontSize(size);
        return this;
    }

    public WMXCell<T> withBackgroundColor(final IndexedColors color) {
        this.style.setBackgroundColor(color);
        return this;
    }

    public WMXCell<T> withComment(final String comment) {
        this.comment = comment;
        return this;
    }

    public String getComment() {
        return comment;
    }

    public WMXCell<T> withMergedCells(final int size) {
        this.mergedColumns = size;
        return this;
    }

    public int getMergedColumns() {
        return mergedColumns;
    }

    public WMXCell<T> withMergedRows(final int size) {
        this.mergedRows = size;
        return this;
    }

    public int getMergedRows() {
        return mergedRows;
    }

    public WMXCell<T> withDataFormat(final String dataFormat) {
        this.getStyle().setDataFormat(dataFormat);
        return this;
    }

    public WMXCell<T> withDataFormat(final short dataForatBuiltin) {
        this.getStyle().setDataFormatBuiltin(dataForatBuiltin);
        return this;
    }
    public WMXCell<T> withBorderTop(final BorderStyle borderTop){
        this.getStyle().setBorderTop(borderTop);
        return this;
    }
    public WMXCell<T> withHorizontalAligment(final HorizontalAlignment horizontalAlignment){
        this.getStyle().setHorizontalAlignment(horizontalAlignment);
        return this;
    }
    public WMXCell<T> withLink(final String link) {
        this.link = link;
        return this;
    }

    public String getLink() {
        return link;
    }

    public WMXCell<T> withFormula(final String formula) {
        this.formula = formula;
        return this;
    }

    public String getFormula() {
        return this.formula;
    }
}
