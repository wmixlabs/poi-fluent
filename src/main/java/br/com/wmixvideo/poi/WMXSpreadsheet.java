package br.com.wmixvideo.poi;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

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

    public WMXSheet withSheet(String name) {
        final WMXSheet sheet = new WMXSheet(name);
        this.sheets.add(sheet);
        return sheet;
    }

    private Workbook build() {
        final HSSFWorkbook woorkBook = new HSSFWorkbook();
        final Map<Integer, HSSFCellStyle> styles = buildGenerateStyles(woorkBook);
        for (WMXSheet sheet : this.sheets) {
            final HSSFSheet sheetCriado = woorkBook.createSheet(sheet.getName());
            for (WMXRow row : sheet.getRows()) {
                final HSSFRow rowCriada = sheetCriado.createRow(Math.max(sheetCriado.getLastRowNum() + 1, 0));
                int posicaoCelula = 0;
                for (WMXCell cell : row.getCells()) {
                    buildGenerateCell(cell, posicaoCelula, rowCriada, sheetCriado, styles);
                    posicaoCelula = posicaoCelula + Math.max(cell.getMergedColumns() - 1, 0) + 1;
                }
            }

            //Processa agrupamento de linhas
            buildGenerateGroupLines(sheet, sheetCriado);

            if (sheet.isAutoSizeColumns()) {
                for (int indiceColuna = 0; indiceColuna <= sheetCriado.getLastRowNum(); indiceColuna++) {
                    sheetCriado.autoSizeColumn(indiceColuna);
                }
            }
        }
        return woorkBook;
    }

    private void buildGenerateGroupLines(final WMXSheet sheet, final HSSFSheet sheetCriado) {
        String agrupador = null;
        List<List<Integer>> agrupamentosTotais = new ArrayList<>();
        List<Integer> linhasAgrupadasAtual = new ArrayList<>();
        for (int i = 0; i < sheet.getRows().size(); i++) {
            final WMXRow dfRow = sheet.getRows().get(i);
            final String agrupadorLinha = dfRow.getGroup();
            if (agrupador != null && agrupadorLinha != null) {
                if (Objects.equals(agrupador, agrupadorLinha)) {
                    linhasAgrupadasAtual.add(i);
                } else {
                    agrupador = agrupadorLinha;
                    agrupamentosTotais.add(linhasAgrupadasAtual);
                    linhasAgrupadasAtual = new ArrayList<>();
                    linhasAgrupadasAtual.add(i);
                }
            } else if (agrupador == null && agrupadorLinha != null) {
                agrupador = agrupadorLinha;
                linhasAgrupadasAtual = new ArrayList<>();
                linhasAgrupadasAtual.add(i);
            } else if (agrupador != null && agrupadorLinha == null) {
                agrupador = null;
                agrupamentosTotais.add(linhasAgrupadasAtual);
                linhasAgrupadasAtual = new ArrayList<>();
            }
        }
        if (!linhasAgrupadasAtual.isEmpty()) {
            agrupamentosTotais.add(linhasAgrupadasAtual);
        }
        for (List<Integer> agrupamento : agrupamentosTotais) {
            sheetCriado.groupRow(agrupamento.get(0) + 1, agrupamento.get(agrupamento.size()-1));
            sheetCriado.setRowSumsBelow(false);
        }
    }

    private void buildGenerateCell(final WMXCell cell, int posicaoCelula, final HSSFRow row, final HSSFSheet sheet, final Map<Integer, HSSFCellStyle> styles) {
        //Crio celula
        final HSSFCell cellCriada = row.createCell(posicaoCelula);

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
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, lastRow, columnIndex, lastCol));
        }
    }

    private Map<Integer, HSSFCellStyle> buildGenerateStyles(HSSFWorkbook woorkBook) {
        final Set<WMXStyle> styles = this.sheets.stream().map(s -> s.getRows()).flatMap(List::stream).map(r -> r.getCells()).flatMap(List::stream).map(c -> c.getStyle()).distinct().collect(Collectors.toSet());
        final Map<Integer, HSSFCellStyle> stylesCriados = new HashMap<>(styles.size());
        for (WMXStyle dfStyle : styles) {
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

            if (dfStyle.getFont() != null || dfStyle.getFontSize() != null || dfStyle.isFontBold() || dfStyle.getFontColor() != null) {
                final HSSFFont font = woorkBook.createFont();
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

    private static Comment buildGenerateComments(final HSSFCell cell, final String comentario) {
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
