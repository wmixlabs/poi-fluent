package com.fincatto.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.math.BigDecimal;

public class DFCell<T> {
    private T value;
    private String comentario;
    private DFStyle style;

    public DFCell(T value) {
        this.value = value;
        this.style = new DFStyle();
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
}
