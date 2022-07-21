package com.fincatto.poi;

import org.apache.poi.ss.usermodel.Row;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

class DFSpreadsheetTest {

    @Test
    public void teste1() throws Exception {
        DFSpreadsheet spreadsheet = new DFSpreadsheet();
        final DFSheet sheet = spreadsheet.withSheet("Teste");

        final DFRow row = sheet.withRow();
        row.withCell("Teste").title();
        row.withCell("O dia que a terra parou").bold();

        final DFRow rowII = sheet.withRow();
        rowII.withCell("Teste II").title();
        rowII.withCell("O dia que a terra parou II").bold();
        rowII.withCell("Esse e um merge de 3 celulas").withMergedCells(3);
        rowII.withCell("Esse e um merge de 2 celulas e 2 linhas").withMergedRows(2).withMergedCells(2);
        rowII.withCell(BigDecimal.valueOf(50.25)).withComment("Comentario teste");

        spreadsheet.toFile("/tmp/planilha"+ LocalDateTime.now().format(DateTimeFormatter.ISO_LOCAL_TIME) +".xls");
    }

//    @Test
//    public void testFreeze() throws Exception {
//        try (DFSpreadsheet spreadsheet = new DFSpreadsheet()) {
//            final DFSheet sheet = spreadsheet.withSheet("Teste");
//            sheet.withRow(0).withCell(0).withValue("Titulo 1");
//            sheet.withRow(5).withCell(10).withValue("O dia em que a terra parou");
//            sheet.freeze(1, 3);
//            spreadsheet.toFile("/tmp/planilha.xlsx");
//        }
//    }
//
//
//    @Test
//    public void testeValue() throws Exception {
//        try (DFSpreadsheet spreadsheet = new DFSpreadsheet()) {
//            final DFSheet sheet = spreadsheet.withSheet("Teste");
////            sheet.withRow(0).withCell(0).withValue("Titulo 1");
//            sheet.withRow(1).withCell().withValue(BigDecimal.TEN);
////            sheet.withRow(5).withCell(10).withValue("O dia em que a terra parou");
////            sheet.freeze(1, 3);
////            spreadsheet.toFile("/tmp/planilha.xlsx");
//
//            Row row = spreadsheet.getRow();
//        }
//    }
}