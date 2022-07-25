package br.com.wmixvideo.poi;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Objects;

public class WMXStyle {

    private BorderStyle borderBottom, borderTop, borderLeft, borderRigth;
    private HorizontalAlignment horizontalAlignment;
    private VerticalAlignment verticalAlignment;
    private String font;
    private Short fontSize;
    private boolean fontBold;
    private IndexedColors backgroundColor, fontColor;
    private Short dataFormatBuiltin;
    private String dataFormat;

    public BorderStyle getBorderBottom() {
        return borderBottom;
    }

    public WMXStyle setBorderBottom(BorderStyle borderBottom) {
        this.borderBottom = borderBottom;
        return this;
    }

    public BorderStyle getBorderTop() {
        return borderTop;
    }

    public WMXStyle setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
        return this;
    }

    public BorderStyle getBorderLeft() {
        return borderLeft;
    }

    public WMXStyle setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
        return this;
    }

    public BorderStyle getBorderRigth() {
        return borderRigth;
    }

    public WMXStyle setBorderRigth(BorderStyle borderRigth) {
        this.borderRigth = borderRigth;
        return this;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public WMXStyle setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public WMXStyle setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }

    public String getFont() {
        return font;
    }

    public WMXStyle setFont(String font) {
        this.font = font;
        return this;
    }

    public Short getFontSize() {
        return fontSize;
    }

    public WMXStyle setFontSize(Short fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public boolean isFontBold() {
        return fontBold;
    }

    public WMXStyle setFontBold(boolean fontBold) {
        this.fontBold = fontBold;
        return this;
    }

    public IndexedColors getBackgroundColor() {
        return backgroundColor;
    }

    public WMXStyle setBackgroundColor(IndexedColors backgroundColor) {
        this.backgroundColor = backgroundColor;
        return this;
    }

    public Short getDataFormatBuiltin() {
        return dataFormatBuiltin;
    }

    public WMXStyle setDataFormatBuiltin(Short dataFormatBuiltin) {
        this.dataFormatBuiltin = dataFormatBuiltin;
        return this;
    }

    public String getDataFormat() {
        return dataFormat;
    }

    public WMXStyle setDataFormat(String dataFormat) {
        this.dataFormat = dataFormat;
        return this;
    }

    public IndexedColors getFontColor() {
        return fontColor;
    }

    public WMXStyle setFontColor(IndexedColors fontColor) {
        this.fontColor = fontColor;
        return this;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        WMXStyle dfStyle = (WMXStyle) o;
        return fontBold == dfStyle.fontBold && borderBottom == dfStyle.borderBottom && borderTop == dfStyle.borderTop && borderLeft == dfStyle.borderLeft && borderRigth == dfStyle.borderRigth && horizontalAlignment == dfStyle.horizontalAlignment && verticalAlignment == dfStyle.verticalAlignment && Objects.equals(font, dfStyle.font) && Objects.equals(fontSize, dfStyle.fontSize) && backgroundColor == dfStyle.backgroundColor && fontColor == dfStyle.fontColor && Objects.equals(dataFormatBuiltin, dfStyle.dataFormatBuiltin) && Objects.equals(dataFormat, dfStyle.dataFormat);
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
        final int hashFontColor = this.fontColor != null ? this.fontColor.hashCode() : 0;
        final int hashBackgroundColor = this.backgroundColor != null ? this.backgroundColor.hashCode() : 0;
        final int hashDataFormatBuiltin = this.dataFormatBuiltin != null ? this.dataFormatBuiltin.hashCode() : 0;
        final int hashDataFormat = this.dataFormat != null ? this.dataFormat.hashCode() : 0;
        return Objects.hash(this.fontBold,
                hashBorderBottom,
                hashBorderTop,
                hashBorderLeft,
                hashBorderRigth,
                hashHorizontalAlignment,
                hashVerticalAlignment,
                hashFont,
                hashFontSize,
                hashFontColor,
                hashBackgroundColor,
                hashDataFormatBuiltin,
                hashDataFormat
        );
    }
}
