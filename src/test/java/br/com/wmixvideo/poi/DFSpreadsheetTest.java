package br.com.wmixvideo.poi;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

class DFSpreadsheetTest {

    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

    @Test
    @Disabled
    public void testeBasico() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");

        final WMXRow row = sheet.withRow();
        row.withCell("Teste").title();
        row.withCell("O dia que a terra parou").bold();
        row.withCell("Esse campo \u00E9 azul").withBackgroundColor(IndexedColors.BLUE);
        row.withCell("Essa letra \u00E9 vermelha").withFontColor(IndexedColors.RED);
        row.withCell("Font Alef").withFontFamily("alef");
        row.withCell("Font 20").withFontSize((short) 20);
        spreadsheet.toFile("/tmp/planilha_basica_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeMerges() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        final WMXRow row = sheet.withRow();
        row.withCell("Teste II").title();
        row.withCell("O dia que a terra parou II").bold();
        row.withCell("Esse e um merge de 3 celulas").withMergedCells(3);
        row.withCell("Esse e um merge de 2 celulas e 2 linhas").withMergedRows(2).withMergedCells(2);
        row.withCell(BigDecimal.valueOf(50.25)).withComment("Comentario teste");
        final WMXRow rowIII = sheet.withRow();
        rowIII.withCell("Teste III").title();
        for (int i = 0; i < 10; i++) {
            rowIII.withCell("Celula " + i);
        }
        spreadsheet.toFile("/tmp/planilha_merges_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeFormatacao() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell(BigDecimal.TEN);
        sheet.withRow().withCell(BigDecimal.valueOf(120.25)).withDataFormat("#,##0.00");
        sheet.withRow().withCell(BigDecimal.valueOf(210.950)).withDataFormat((short) 5);
        sheet.withRow().withCell(BigDecimal.valueOf(1120.10)).withDataFormat((short) 0x28);
        sheet.withRow().withCell(LocalDate.now()).withDataFormat((short) 0xe);
        sheet.withRow().withCell(LocalDate.now()).withDataFormat("dd/MM/YYYY");
        sheet.withRow().withCell(LocalDate.now()).withDataFormat("dd/MM/YY");
        sheet.withRow().withCell(LocalDateTime.now()).withDataFormat("dd/MM/YYYY hh:mm:ss");
        sheet.withRow().withCell(LocalDateTime.now()).withDataFormat("hh:mm:ss");
        sheet.withRow().withCell(LocalDateTime.now()).withDataFormat((short) 0x16);
        sheet.withRow().withCell(Boolean.TRUE);
        sheet.withRow().withCell(Boolean.FALSE);
        spreadsheet.toFile("/tmp/planilha_formatos_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeLink() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell("Filme");
        sheet.withRow().withCell("tt0899128").withLink("https://www.imdb.com/title/tt0899128/?ref_=nv_sr_srsg_0");
        spreadsheet.toFile("/tmp/planilha_link_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeAutoSize() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        final WMXRow dfRow = sheet.withRow();
        dfRow.withCell("Este \u00E9 um texto longo");
        dfRow.withCell("Texto");
        sheet.withRow().withCell("curto");
        sheet.withAutoSizeColumns(true);
        spreadsheet.toFile("/tmp/planilha_autosize_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeFormula() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell("").withFormula("DATE(2020,12,1)");
        spreadsheet.toFile("/tmp/planilha_formula_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeAgrupamento() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");

        final WMXRow dfRow = sheet.withRow().withGroup("Agrupador1");
        dfRow.withCell("Linha 1 agrupada").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRow.withCell("Celula 1").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRow.withCell("Celula 2").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRow.withCell("Celula 3").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRow.withCell("Celula 4").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRow.withCell("Celula 5").withBackgroundColor(IndexedColors.GREY_25_PERCENT);

        final WMXRow dfRowII = sheet.withRow().withGroup("Agrupador1");
        dfRowII.withCell("Linha 2 agrupada");
        dfRowII.withCell("Celula 1");
        dfRowII.withCell("Celula 2");
        dfRowII.withCell("Celula 3");
        dfRowII.withCell("Celula 4");
        dfRowII.withCell("Celula 5");

        final WMXRow dfRowIII = sheet.withRow().withGroup("Agrupador1");
        dfRowIII.withCell("Linha 3 agrupada");
        dfRowIII.withCell("Celula 1");
        dfRowIII.withCell("Celula 2");
        dfRowIII.withCell("Celula 3");
        dfRowIII.withCell("Celula 4");
        dfRowIII.withCell("Celula 5");

        final WMXRow dfRowIV = sheet.withRow().withGroup("Agrupador2");
        dfRowIV.withCell("Linha 4 agrupada").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRowIV.withCell("Celula 1").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRowIV.withCell("Celula 2").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRowIV.withCell("Celula 3").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRowIV.withCell("Celula 4").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        dfRowIV.withCell("Celula 5").withBackgroundColor(IndexedColors.GREY_25_PERCENT);

        final WMXRow dfRowV = sheet.withRow().withGroup("Agrupador2");
        dfRowV.withCell("Linha 5 agrupada");
        dfRowV.withCell("Celula 1");
        dfRowV.withCell("Celula 2");
        dfRowV.withCell("Celula 3");
        dfRowV.withCell("Celula 4");
        dfRowV.withCell("Celula 5");

        final WMXRow dfRowVI = sheet.withRow();
        dfRowVI.withCell("Linha 6 desagrupada");
        dfRowVI.withCell("Celula 1");
        dfRowVI.withCell("Celula 2");
        dfRowVI.withCell("Celula 3");
        dfRowVI.withCell("Celula 4");
        dfRowVI.withCell("Celula 5");

        spreadsheet.toFile("/tmp/planilha_agrupamento_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }
}