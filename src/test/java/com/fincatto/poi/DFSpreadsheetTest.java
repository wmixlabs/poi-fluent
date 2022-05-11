package com.fincatto.poi;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.IOException;

class DFSpreadsheetTest {

    @Test
    public void teste1() throws Exception {
        try (DFSpreadsheet spreadsheet = new DFSpreadsheet()) {
            Assertions.assertNotNull(spreadsheet
                    .withSheet("Teste")
                    .withRow(5)
                    .withCell(10)
                    .withValue("O dia em que a terra parou")

            );
            spreadsheet.toFile("C:\\Users\\diego\\Downloads\\planilha.xlsx");
        }
    }
}