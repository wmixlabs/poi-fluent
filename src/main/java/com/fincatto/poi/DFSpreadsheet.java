package com.fincatto.poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

public class DFSpreadsheet {

    //    private final Workbook workbook;
    private List<DFSheet> sheets;

    public DFSpreadsheet() {
        this.sheets = new ArrayList<>();
    }

    public DFSheet withSheet(String name) {
        final DFSheet sheet = new DFSheet(name);
        this.sheets.add(sheet);
        return sheet;
    }

    private Workbook build() {
        final HSSFWorkbook woorkBook = new HSSFWorkbook();
        final Map<Integer, HSSFCellStyle> styles = gerarStyles(woorkBook);
        for (DFSheet sheet : this.sheets) {
            final HSSFSheet sheetCriado = woorkBook.createSheet(sheet.getName());
            for (DFRow row : sheet.getRows()) {
                final HSSFRow rowCriada = sheetCriado.createRow(Math.max(sheetCriado.getLastRowNum() + 1, 0));
                for (DFCell cell : row.getCells()) {
                    final HSSFCell cellCriada = rowCriada.createCell(Math.max(rowCriada.getLastCellNum(), 0));
                    cellCriada.setCellStyle(styles.get(cell.getStyle().hashCode()));
                    final Object value = cell.getValue();
                    if (value != null) {
                        if (value instanceof String)
                            cellCriada.setCellValue(value.toString());
                        else if (value instanceof BigDecimal) {
                            cellCriada.setCellValue(((BigDecimal) value).doubleValue());
                        } else if (value instanceof Number) {
                            cellCriada.setCellValue(((Number) value).doubleValue());
                        } else {
                            cellCriada.setCellValue(value.toString());
                        }
                    } else {
                        cellCriada.setCellValue("");
                    }
                }
            }
        }
        return woorkBook;
    }

    private Map<Integer, HSSFCellStyle> gerarStyles(HSSFWorkbook woorkBook) {
        final Set<DFStyle> styles = this.sheets.stream().map(s -> s.getRows()).flatMap(List::stream).map(r -> r.getCells()).flatMap(List::stream).map(c -> c.getStyle()).distinct().collect(Collectors.toSet());
        final Map<Integer, HSSFCellStyle> stylesCriados = new HashMap<>(styles.size());
        for (DFStyle dfStyle : styles) {
            final HSSFCellStyle cellStyle = woorkBook.createCellStyle();
            if (dfStyle.getBackgroundColor() != null) {
                cellStyle.setFillForegroundColor(dfStyle.getBackgroundColor().getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            if (dfStyle.getBorderBottom() != null) {
                cellStyle.setBorderBottom(dfStyle.getBorderBottom());
            }

            if (dfStyle.getBorderTop() != null) {
                cellStyle.setBorderTop(dfStyle.getBorderTop());
            }

            if (dfStyle.getBorderLeft() != null) {
                cellStyle.setBorderLeft(dfStyle.getBorderLeft());
            }

            if (dfStyle.getBorderRigth() != null) {
                cellStyle.setBorderRight(dfStyle.getBorderRigth());
            }
            if(dfStyle.getFont() != null || dfStyle.getFontSize() != null || dfStyle.isFontBold()){
                final HSSFFont font = woorkBook.createFont();
                font.setBold(dfStyle.isFontBold());
                if(dfStyle.getFont() != null){
                    font.setFontName(dfStyle.getFont());
                }
                if(dfStyle.getFontSize() != null){
                    font.setFontHeightInPoints(dfStyle.getFontSize());
                }
                cellStyle.setFont(font);
            }
            stylesCriados.put(dfStyle.hashCode(), cellStyle);
        }

        return stylesCriados;

    }

    public void toFile(final String path) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(path)) {
            try (Workbook workbook = build()) {
                workbook.write(outputStream);
            }
        }
    }

    public byte[] toByteArray() throws IOException {
        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            try (Workbook workbook = build()) {
                workbook.write(byteArrayOutputStream);
                return byteArrayOutputStream.toByteArray();
            }
        }
    }
}
