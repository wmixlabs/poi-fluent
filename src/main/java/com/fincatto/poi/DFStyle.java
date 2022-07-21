package com.fincatto.poi;

import org.apache.poi.ss.usermodel.*;

import java.util.Objects;

public class DFStyle {

    private BorderStyle borderBottom, borderTop, borderLeft, borderRigth;
    private HorizontalAlignment horizontalAlignment;
    private VerticalAlignment verticalAlignment;
    private String font;
    private Short fontSize;
    private boolean fontBold;
    private IndexedColors backgroundColor;

    public BorderStyle getBorderBottom() {
        return borderBottom;
    }

    public DFStyle setBorderBottom(BorderStyle borderBottom) {
        this.borderBottom = borderBottom;
        return this;
    }

    public BorderStyle getBorderTop() {
        return borderTop;
    }

    public DFStyle setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
        return this;
    }

    public BorderStyle getBorderLeft() {
        return borderLeft;
    }

    public DFStyle setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
        return this;
    }

    public BorderStyle getBorderRigth() {
        return borderRigth;
    }

    public DFStyle setBorderRigth(BorderStyle borderRigth) {
        this.borderRigth = borderRigth;
        return this;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public DFStyle setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public DFStyle setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }

    public String getFont() {
        return font;
    }

    public DFStyle setFont(String font) {
        this.font = font;
        return this;
    }

    public Short getFontSize() {
        return fontSize;
    }

    public DFStyle setFontSize(Short fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public boolean isFontBold() {
        return fontBold;
    }

    public DFStyle setFontBold(boolean fontBold) {
        this.fontBold = fontBold;
        return this;
    }

    public IndexedColors getBackgroundColor() {
        return backgroundColor;
    }

    public DFStyle setBackgroundColor(IndexedColors backgroundColor) {
        this.backgroundColor = backgroundColor;
        return this;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        DFStyle dfStyle = (DFStyle) o;
        return fontBold == dfStyle.fontBold && borderBottom == dfStyle.borderBottom && borderTop == dfStyle.borderTop && borderLeft == dfStyle.borderLeft && borderRigth == dfStyle.borderRigth && horizontalAlignment == dfStyle.horizontalAlignment && verticalAlignment == dfStyle.verticalAlignment && Objects.equals(font, dfStyle.font) && Objects.equals(fontSize, dfStyle.fontSize) && backgroundColor == dfStyle.backgroundColor;
    }

    @Override
    public int hashCode() {
        final int hashBorderBottom = this.borderBottom != null ? this.borderBottom.hashCode() : 0;
        final int hashBorderTop = this.borderTop != null ? this.borderTop.hashCode() : 0;
        final int hashBorderLeft = this.borderLeft != null ? this.borderLeft.hashCode() : 0;
        final int hashBorderRigth = this.borderRigth != null ? this.borderRigth.hashCode() : 0;
        final int hashHorizontalAlignment = this.horizontalAlignment != null ? this.horizontalAlignment.hashCode() : 0;
        final int hashVerticalAlignment = this.verticalAlignment != null ? this.verticalAlignment.hashCode() : 0;
        final int hashFont = this.font != null ? this.font.hashCode() : 0;
        final int hashFontSize = this.fontSize != null ? this.fontSize.hashCode() : 0;
        final int hashBackgroundColor = this.backgroundColor != null ? this.backgroundColor.hashCode() : 0;
        return Objects.hash(this.fontBold,
                hashBorderBottom,
                hashBorderTop,
                hashBorderLeft,
                hashBorderRigth,
                hashHorizontalAlignment,
                hashVerticalAlignment,
                hashFont,
                hashFontSize,
                hashBackgroundColor);
    }
}
