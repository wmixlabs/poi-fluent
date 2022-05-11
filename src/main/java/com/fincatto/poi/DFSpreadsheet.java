package com.fincatto.poi;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class DFSpreadsheet implements AutoCloseable {

    private final Workbook workbook;

    public DFSpreadsheet() {
        this.workbook = new XSSFWorkbook();
    }

    public DFSheet withSheet(String name) {
        final Sheet sheet = workbook.getSheet(name);
        if (sheet == null) {
            return new DFSheet(workbook.createSheet(name));
        }
        return new DFSheet(sheet);
    }

    public void toFile(final String path) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(path)) {
            workbook.write(outputStream);
        }
    }

    @Override
    public void close() throws Exception {
        workbook.close();
    }
}
