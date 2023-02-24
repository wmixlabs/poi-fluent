package br.com.wmixvideo.poi;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

public class WMXSpreadsheet {

    private final List<WMXSheet> sheets;

    public WMXSpreadsheet() {
        this.sheets = new ArrayList<>();
    }

    public WMXSheet withSheet(final String name) {
        final WMXSheet sheet = new WMXSheet(name, this);
        this.sheets.add(sheet);
        return sheet;
    }

    private Workbook build(WMXFormat format) {
        final Workbook woorkBook = WMXFormat.XLS.equals(format) ? new HSSFWorkbook() : new XSSFWorkbook();
        final Map<Integer, CellStyle> styles = buildGenerateStyles(woorkBook);
        for (WMXSheet sheet : this.sheets) {
            final Set<Integer> columnsHidden = new HashSet<>();
            final Sheet sheetCriado = woorkBook.createSheet(sheet.getName());
            for (WMXRow row : sheet.getRows()) {
                final Row rowCriada = sheetCriado.createRow(Math.max(sheetCriado.getLastRowNum() + 1, 0));
                int posicaoCelula = 0;

                sheetCriado.getRow(rowCriada.getRowNum()).setZeroHeight(row.isHiddenRow());
                for (WMXCell<?> cell : row.getCells()) {
                    if (cell.isHiddenColumn()) {
                        columnsHidden.add(cell.getIndex() - 1);
                    }
                    buildGenerateCell(cell, posicaoCelula, rowCriada, sheetCriado, styles);
                    posicaoCelula = posicaoCelula + Math.max(cell.getMergedColumns() - 1, 0) + 1;

                }
            }
            for (Integer indexColumn : columnsHidden) {
                sheetCriado.setColumnHidden(indexColumn, true);
            }
            sheetCriado.createFreezePane(sheet.getFreezeCols(), sheet.getFreezeRows());

            //Processa agrupamento de linhas
            buildGenerateAssociedLines(sheet, sheetCriado);

            if (sheet.isAutoSizeColumns()) {
                for (int indiceColuna = 0; indiceColuna <= sheetCriado.getLastRowNum(); indiceColuna++) {
                    sheetCriado.autoSizeColumn(indiceColuna);
                }
            }

            if (sheet.getAutoFilterRange() != null) {
                sheetCriado.setAutoFilter(new CellRangeAddress(sheet.getAutoFilterRange().getFirstRow(), sheet.getAutoFilterRange().getLastRow() - 1,
                        sheet.getAutoFilterRange().getFirstColumn(), sheet.getAutoFilterRange().getLastColumn() - 1));
            }
        }
        return woorkBook;
    }


    private void buildGenerateAssociedLines(final WMXSheet sheet, final Sheet sheetCriado) {
        final List<Pair<Integer, Integer>> grupos = new ArrayList<>();
        for (WMXRow row : sheet.getRows()) {
            if (!row.getGroupedRows().isEmpty()) {
                System.out.printf("Grupo criado: %s, %s\n", row.getIndex(), row.getGroupedRows().get(row.getGroupedRows().size() - 1).getIndex() - 1);
                sheetCriado.groupRow(row.getIndex(), row.getGroupedRows().get(row.getGroupedRows().size() - 1).getIndex() - 1);
                sheetCriado.setRowSumsBelow(false);
            }
        }
    }

    private void buildGenerateCell(final WMXCell<?> cell, int posicaoCelula, final Row row, final Sheet sheet, final Map<Integer, CellStyle> styles) {
        //Crio celula
        final Cell cellCriada = row.createCell(posicaoCelula);

        //Formato celula
        cellCriada.setCellStyle(styles.get(cell.getStyle().hashCode()));

        //Preencho valor da celula
        final Object value = cell.getValue();
        if (value != null) {
            if (value instanceof String)
                cellCriada.setCellValue(value.toString());
            else if (value instanceof BigDecimal) {
                cellCriada.setCellValue(((BigDecimal) value).doubleValue());
            } else if (value instanceof Number) {
                cellCriada.setCellValue(((Number) value).doubleValue());
            } else if (value instanceof LocalDate) {
                cellCriada.setCellValue(((LocalDate) value));
            } else if (value instanceof LocalDateTime) {
                cellCriada.setCellValue(((LocalDateTime) value));
            } else if (value instanceof Boolean) {
                cellCriada.setCellValue(((Boolean) value));
            } else {
                cellCriada.setCellValue(value.toString());
            }
        } else {
            cellCriada.setCellValue("");
        }

        //Crio comentario na celula
        if (cell.getComment() != null) {
            cellCriada.setCellComment(buildGenerateComments(cellCriada, cell.getComment()));
        }

        //Preencho formula da celula
        if (cell.getFormula() != null) {
            cellCriada.setCellFormula(cell.getFormula());
        }

        //Crio link na celula
        if (cell.getLink() != null) {
            final Hyperlink hyperlink = row.getSheet().getWorkbook().getCreationHelper().createHyperlink(HyperlinkType.URL);
            hyperlink.setAddress(cell.getLink());
            cellCriada.setHyperlink(hyperlink);
        }

        //Crio regiao com merge
        if (cell.getMergedColumns() > 0 || cell.getMergedRows() > 0) {
            final int rowIndex = cellCriada.getRowIndex();
            final int lastRow = cell.getMergedRows() > 0 ? (cellCriada.getRowIndex() + cell.getMergedRows()) - 1 : cellCriada.getRowIndex();
            final int columnIndex = cellCriada.getColumnIndex();
            final int lastCol = cell.getMergedColumns() > 0 ? (cellCriada.getColumnIndex() + cell.getMergedColumns()) - 1 : cellCriada.getColumnIndex();

            final CellRangeAddress region = new CellRangeAddress(rowIndex, lastRow, columnIndex, lastCol);
            sheet.addMergedRegion(region);

            if (cell.getStyle().getBorderTop() != null) {
                RegionUtil.setBorderTop(cell.getStyle().getBorderTop(), region, sheet);
            }

            if (cell.getStyle().getBorderBottom() != null) {
                RegionUtil.setBorderBottom(cell.getStyle().getBorderBottom(), region, sheet);
            }

            if (cell.getStyle().getBorderLeft() != null) {
                RegionUtil.setBorderLeft(cell.getStyle().getBorderLeft(), region, sheet);
            }

            if (cell.getStyle().getBorderRight() != null) {
                RegionUtil.setBorderRight(cell.getStyle().getBorderRight(), region, sheet);
            }
        }
    }

    private Map<Integer, CellStyle> buildGenerateStyles(Workbook woorkBook) {
        final Set<WMXStyle> styles = this.sheets.stream().map(WMXSheet::getRows).flatMap(List::stream).map(WMXRow::getCells).flatMap(List::stream).map(WMXCell::getStyle).collect(Collectors.toSet());
        final Map<Integer, CellStyle> stylesCriados = new HashMap<>(styles.size());
        for (WMXStyle dfStyle : styles) {
            final CellStyle cellStyle = woorkBook.createCellStyle();

            if (dfStyle.getHorizontalAlignment() != null) {
                cellStyle.setAlignment(dfStyle.getHorizontalAlignment());
            }

            if (dfStyle.getBackgroundColor() != null) {
                cellStyle.setFillForegroundColor(dfStyle.getBackgroundColor().getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            if (dfStyle.getCustomBackgroundColor() != null && woorkBook instanceof XSSFWorkbook) {
                ((XSSFCellStyle) cellStyle).setFillForegroundColor(new XSSFColor(dfStyle.getCustomBackgroundColor(), ((XSSFWorkbook) woorkBook).getStylesSource().getIndexedColors()));
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

            if (dfStyle.getBorderRight() != null) {
                cellStyle.setBorderRight(dfStyle.getBorderRight());
            }

            if (dfStyle.getFont() != null || dfStyle.getFontSize() != null || dfStyle.isFontBold() || dfStyle.getFontColor() != null || dfStyle.getCustomFontColor() != null) {
                final Font font = woorkBook.createFont();
                font.setBold(dfStyle.isFontBold());
                if (dfStyle.getFont() != null) {
                    font.setFontName(dfStyle.getFont());
                }
                if (dfStyle.getFontSize() != null) {
                    font.setFontHeightInPoints(dfStyle.getFontSize());
                }
                if (dfStyle.getFontColor() != null) {
                    font.setColor(dfStyle.getFontColor().getIndex());
                }
                if (dfStyle.getCustomFontColor() != null && woorkBook instanceof XSSFWorkbook) {
                    ((XSSFFont) font).setColor(new XSSFColor(dfStyle.getCustomFontColor(), ((XSSFWorkbook) woorkBook).getStylesSource().getIndexedColors()));
                }
                cellStyle.setFont(font);
            }

            if (dfStyle.getDataFormatBuiltin() != null) {
                cellStyle.setDataFormat(dfStyle.getDataFormatBuiltin());
            }

            if (dfStyle.getDataFormat() != null) {
                cellStyle.setDataFormat(woorkBook.createDataFormat().getFormat(dfStyle.getDataFormat()));
            }

            stylesCriados.put(dfStyle.hashCode(), cellStyle);
        }

        return stylesCriados;

    }

    private static Comment buildGenerateComments(final Cell cell, final String comentario) {
        if (comentario != null && !comentario.isBlank()) {
            final CreationHelper factory = cell.getRow().getSheet().getWorkbook().getCreationHelper();

            final ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex());
            anchor.setCol2(cell.getColumnIndex() + 3);
            anchor.setRow1(cell.getRowIndex());
            anchor.setRow2(cell.getRowIndex() + 4);

            final Comment comment = cell.getSheet().createDrawingPatriarch().createCellComment(anchor);
            comment.setString(factory.createRichTextString(comentario));
            return comment;
        }
        return null;
    }

    public void toFile(final String path) throws IOException {
        toFile(WMXFormat.XLSX, path);
    }

    public void toFile(final WMXFormat format, final String path) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(path)) {
            try (Workbook workbook = build(format)) {
                workbook.write(outputStream);
            }
        }
    }

    public byte[] toByteArray() throws IOException {
        return toByteArray(WMXFormat.XLSX);
    }

    public byte[] toByteArray(final WMXFormat format) throws IOException {
        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            try (Workbook workbook = build(format)) {
                workbook.write(byteArrayOutputStream);
                return byteArrayOutputStream.toByteArray();
            }
        }
    }
}
