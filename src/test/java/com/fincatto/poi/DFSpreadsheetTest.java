package com.fincatto.poi;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

class DFSpreadsheetTest {

    @Test
    public void teste1() throws Exception {
        try (DFSpreadsheet spreadsheet = new DFSpreadsheet()) {
            Assertions.assertNotNull(spreadsheet.withSheet("Teste").withRow(5).withCell(10).withValue("O dia em que a terra parou"));
            spreadsheet.toFile("C:\\Users\\diego\\Downloads\\planilha.xlsx");
        }
    }

    @Test
    public void testFreeze() throws Exception {
        try (DFSpreadsheet spreadsheet = new DFSpreadsheet()) {
            final DFSheet sheet = spreadsheet.withSheet("Teste");
            sheet.withRow(0).withCell(0).withValue("Titulo 1");
            sheet.withRow(5).withCell(10).withValue("O dia em que a terra parou");
            sheet.freeze(1, 3).unfreeze();
            spreadsheet.toFile("C:\\Users\\diego\\Downloads\\planilha.xlsx");
        }
    }
}