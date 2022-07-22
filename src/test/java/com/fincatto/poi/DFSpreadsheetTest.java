package com.fincatto.poi;

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
        final DFSpreadsheet spreadsheet = new DFSpreadsheet();
        final DFSheet sheet = spreadsheet.withSheet("Teste");

        final DFRow row = sheet.withRow();
        row.withCell("Teste").title();
        row.withCell("O dia que a terra parou").bold();
        row.withCell("Esse campo \u00E9 azul").withBackgroundColor(IndexedColors.BLUE);
        row.withCell("Essa letra \u00E9 vermelha").withFontColor(IndexedColors.RED);
        row.withCell("Font Alef").withFontFamily("alef");
        row.withCell("Font 20").withFontSize((short) 20);
        spreadsheet.toFile("/tmp/planilha_basica"+ LocalDateTime.now().format(FORMATTER) +".xls");
    }

    @Test
    public void testeMerges() throws Exception {
        final DFSpreadsheet spreadsheet = new DFSpreadsheet();
        final DFSheet sheet = spreadsheet.withSheet("Teste");
        final DFRow row = sheet.withRow();
        row.withCell("Teste II").title();
        row.withCell("O dia que a terra parou II").bold();
        row.withCell("Esse e um merge de 3 celulas").withMergedCells(3);
        row.withCell("Esse e um merge de 2 celulas e 2 linhas").withMergedRows(2).withMergedCells(2);
        row.withCell(BigDecimal.valueOf(50.25)).withComment("Comentario teste");
        final DFRow rowIII = sheet.withRow();
        rowIII.withCell("Teste III").title();
        for (int i = 0; i < 10; i++) {
            rowIII.withCell("Celula "+i);
        }
        spreadsheet.toFile("/tmp/planilha_merges"+ LocalDateTime.now().format(FORMATTER) +".xls");
    }

    @Test
    public void testeFormatacao() throws Exception {
        final DFSpreadsheet spreadsheet = new DFSpreadsheet();
        final DFSheet sheet = spreadsheet.withSheet("Teste");
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
        spreadsheet.toFile("/tmp/planilha_formatos"+ LocalDateTime.now().format(FORMATTER) +".xls");
    }

    @Test
    public void testeLink() throws Exception {
        final DFSpreadsheet spreadsheet = new DFSpreadsheet();
        final DFSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell("Filme");
        sheet.withRow().withCell("tt0899128").withLink("https://www.imdb.com/title/tt0899128/?ref_=nv_sr_srsg_0");
        spreadsheet.toFile("/tmp/planilha_link"+ LocalDateTime.now().format(FORMATTER) +".xls");
    }

    @Test
    public void testeAutoSize() throws Exception {
        final DFSpreadsheet spreadsheet = new DFSpreadsheet();
        final DFSheet sheet = spreadsheet.withSheet("Teste");
        final DFRow dfRow = sheet.withRow();
        dfRow.withCell("Este \u00E9 um texto longo");
        dfRow.withCell("Texto");
        sheet.withRow().withCell("curto");
        sheet.withAutoSizeColumns(true);
        spreadsheet.toFile("/tmp/planilha_autosize"+ LocalDateTime.now().format(FORMATTER) +".xls");
    }

    @Test
    public void testeFormula() throws Exception {
        final DFSpreadsheet spreadsheet = new DFSpreadsheet();
        final DFSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell("").withFormula("DATE(2020,12,1)");
        spreadsheet.toFile("/tmp/planilha_formula"+ LocalDateTime.now().format(FORMATTER) +".xls");
    }
}