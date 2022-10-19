package br.com.wmixvideo.poi;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

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
        spreadsheet.toFile(WMXFormat.XLS, "/tmp/planilha_basica_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeMerges() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        final WMXRow row = sheet.withRow();
        row.withCell("Teste II").title();
        row.withCell("O dia que a terra parou II").bold();
        row.withCell("Esse e um merge de 3 celulas").withMergedColumns(3);
        row.withCell("Esse e um merge de 2 celulas e 2 linhas").withMergedRows(2).withMergedColumns(2);
        row.withCell(BigDecimal.valueOf(50.25)).withComment("Comentario teste");
        final WMXRow rowIII = sheet.withRow();
        rowIII.withCell("Teste III").title();
        for (int i = 0; i < 10; i++) {
            rowIII.withCell("Celula " + i);
        }
        spreadsheet.toFile("/tmp/planilha_merges_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
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
        spreadsheet.toFile("/tmp/planilha_formatos_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeLink() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell("Filme");
        sheet.withRow().withCell("tt0899128").withLink("https://www.imdb.com/title/tt0899128/?ref_=nv_sr_srsg_0");
        spreadsheet.toFile("/tmp/planilha_link_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
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
        spreadsheet.toFile("/tmp/planilha_autosize_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeFormula() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell("").withFormula("DATE(2020,12,1)");
        spreadsheet.toFile("/tmp/planilha_formula_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeAgrupamento() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");

        sheet.withRow().withGroup("Agrupador1")
                .withCell("Linha 1 agrupada").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 1").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 2").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 3").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 4").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.GREY_25_PERCENT);

        sheet.withRow().withGroup("Agrupador1")
                .withCell("Linha 2 agrupada").and()
                .withCell("Celula 1").and()
                .withCell("Celula 2").and()
                .withCell("Celula 3").and()
                .withCell("Celula 4").and()
                .withCell("Celula 5");

        sheet.withRow().withGroup("Agrupador1")
                .withCell("Linha 3 agrupada").and()
                .withCell("Celula 1").and()
                .withCell("Celula 2").and()
                .withCell("Celula 3").and()
                .withCell("Celula 4").and()
                .withCell("Celula 5");

        sheet.withRow().withGroup("Agrupador2")
                .withCell("Linha 4 agrupada").withBackgroundColor(IndexedColors.GREY_50_PERCENT).and()
                .withCell("Celula 1").withBackgroundColor(IndexedColors.GREY_50_PERCENT).and()
                .withCell("Celula 2").withBackgroundColor(IndexedColors.GREY_50_PERCENT).and()
                .withCell("Celula 3").withBackgroundColor(IndexedColors.GREY_50_PERCENT).and()
                .withCell("Celula 4").withBackgroundColor(IndexedColors.GREY_50_PERCENT).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.GREY_50_PERCENT);

        sheet.withRow().withGroup("Agrupador2")
                .withCell("Linha 5 agrupada").and()
                .withCell("Celula 1").and()
                .withCell("Celula 2").and()
                .withCell("Celula 3").and()
                .withCell("Celula 4").and()
                .withCell("Celula 5");

        sheet.withRow()
                .withCell("Linha 6 desagrupada").and()
                .withCell("Celula 1").withBackgroundColor(IndexedColors.BLUE).and()
                .withCell("Celula 2").withBackgroundColor(IndexedColors.LIGHT_BLUE).and()
                .withCell("Celula 3").withBackgroundColor(IndexedColors.BLUE1).and()
                .withCell("Celula 4").withBackgroundColor(IndexedColors.CORNFLOWER_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.PALE_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.ROYAL_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.SKY_BLUE);

        spreadsheet.toFile("/tmp/planilha_agrupamento_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeCoresPersonalizadas() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");

        sheet.withRow().withGroup("Agrupador1")
                .withCell("Linha 1 agrupada").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 1").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 2").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 3").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 4").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.GREY_25_PERCENT);

        sheet.withRow().withGroup("Agrupador1")
                .withCell("Linha 2 agrupada").and()
                .withCell("Celula 1").and()
                .withCell("Celula 2").and()
                .withCell("Celula 3").and()
                .withCell("Celula 4").and()
                .withCell("Celula 5");

        sheet.withRow().withGroup("Agrupador1")
                .withCell("Linha 3 agrupada").and()
                .withCell("Celula 1").and()
                .withCell("Celula 2").and()
                .withCell("Celula 3").and()
                .withCell("Celula 4").and()
                .withCell("Celula 5");

        sheet.withRow().withGroup("Agrupador2")
                .withCell("Linha 4 agrupada").withBackgroundColor(IndexedColors.GREY_50_PERCENT).and()
                .withCell("Celula 1").withBackgroundColor(new Color(0xFCCACA)).withFontColor(new Color(0x9B3131)).and()
                .withCell("Celula 2").withBackgroundColor(new Color(0xBBD0AD)).withFontColor(new Color(0x2F6222)).and()
                .withCell("Celula 3").withBackgroundColor(new Color(0xFDF3C2)).withFontColor(new Color(0x726C0B)).and()
                .withCell("Celula 4").withBackgroundColor(new Color(0xE5C9EE)).withFontColor(new Color(0x421652)).and()
                .withCell("Celula 5").withBackgroundColor(new Color(0xAECAF5)).withFontColor(new Color(0x1A4B9B));

        sheet.withRow().withGroup("Agrupador2")
                .withCell("Linha 5 agrupada").and()
                .withCell("Celula 1").and()
                .withCell("Celula 2").and()
                .withCell("Celula 3").and()
                .withCell("Celula 4").and()
                .withCell("Celula 5");

        sheet.withRow()
                .withCell("Linha 6 desagrupada").and()
                .withCell("Celula 1").withBackgroundColor(IndexedColors.BLUE).and()
                .withCell("Celula 2").withBackgroundColor(IndexedColors.LIGHT_BLUE).and()
                .withCell("Celula 3").withBackgroundColor(IndexedColors.BLUE1).and()
                .withCell("Celula 4").withBackgroundColor(IndexedColors.CORNFLOWER_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.PALE_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.ROYAL_BLUE).and()
                .withCell("Celula 5").withBackgroundColor(IndexedColors.SKY_BLUE);

        spreadsheet.toFile("/tmp/planilha_cores_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeOcultaExibicao() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell("Coluna 1 não escondida").and()
                .withCell("Coluna 2 expandida").withMergedColumns(2);
        sheet.withRow().withCell(null).and().withCell("Coluna 2 escondida ").withHiddenColumn(true).and().withCell("Coluna 3 não escondida");
//        sheet.withRow().withHiddenRow(true).withCell("Linha escondida");
        sheet.withAutoSizeColumns(true);
        spreadsheet.toFile("/tmp/planilha_oculta_coluna_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeFormataDataPadrao() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        sheet.withRow().withCell(LocalDate.now()).and()
                .withCell(LocalDateTime.now()).and()
                .withCell(new Date());
        spreadsheet.toFile("/tmp/planilha_data_padrao" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeIndexColumns() {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        Assertions.assertEquals(3, sheet.withRow().withCell("Coluna 1").and()
                .withCell("Coluna 2").and()
                .withCell("Coluna 3").getIndex());

        Assertions.assertEquals(1, sheet.withRow().
                withCell("Coluna 1 merge de 4 celulas").withMergedColumns(4).getIndex());

        Assertions.assertEquals(5, sheet.withRow().
                withCell("Coluna 1 merge de 4 celulas").withMergedColumns(4).and()
                .withCell("Coluna 2 ").getIndex());
    }

    @Test
    @Disabled
    public void testeIndexRows() {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        final WMXRow row1 = sheet.withRow();
        final WMXRow row2 = sheet.withRow();

        Assertions.assertEquals(2, row2.getIndex());
        Assertions.assertEquals(1, row1.getIndex());
    }

    @Test
    @Disabled
    public void testeIndexLetter() {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Teste");

        Assertions.assertEquals("A", sheet.withRow().withCell("Teste").getIndexLetter());
        Assertions.assertEquals("B", sheet.withRow().withCell("Coluna1").and().withCell("Coluna2").getIndexLetter());
    }

    @Test
    @Disabled
    public void testeSubTotal() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Teste");

        sheet.withRow().withCell("COLUNA A").subtitle()
                .and().withCell("COLUNA B").subtitle();

        sheet.withRow().withCell("A").and().withCell(BigDecimal.valueOf(20000)).currency();
        sheet.withRow().withCell("A").and().withCell(BigDecimal.valueOf(20000)).currency();
        sheet.withRow().withCell("B").and().withCell(BigDecimal.valueOf(20000)).currency();
        sheet.withRow().withCell("B").and().withCell(BigDecimal.valueOf(20000)).currency();
        sheet.withRow().withCell(null).and().withCell("Teste");
        sheet.withRow().withCell(null).and().withCell(null).subtotal().currency();
        sheet.withRow().withCell(null).and().withCell(null).subtotal().currency();
        sheet.withRow().withCell(null).and().withCell(null).subtotal().currency();
        sheet.withRow().withCell(null).and().withCell(null).subtotal().currency();

        spreadsheet.toFile("/tmp/planilha_sub_total" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeFiltro() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste").withAutoFilter(0,0, 1, 3);

        sheet.withRow().withCell("COLUNA A").subtitle()
                .and().withCell("COLUNA B").subtitle()
                .and().withCell("COLUNA C").subtitle();

        sheet.withRow().withCell("A").and().withCell(BigDecimal.valueOf(20000)).currency().and().withCell(BigDecimal.valueOf(10000)).currency();
        sheet.withRow().withCell("A").and().withCell(BigDecimal.valueOf(20000)).currency();
        sheet.withRow().withCell("B").and().withCell(BigDecimal.valueOf(20000)).currency();
        sheet.withRow().withCell("B").and().withCell(BigDecimal.valueOf(20000)).currency();
        sheet.withRow().withCell(null).and().withCell("Teste");
        sheet.withRow().withCell(null).and().withCell(null).subtotal().currency();
        sheet.withRow().withCell(null).and().withCell(null).currency();
        sheet.withRow().withCell(Boolean.TRUE).and().withCell(null).and().withCell(null);
        spreadsheet.toFile("/tmp/planilha_auto_filter" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }
    
    @Test
    @Disabled
    public void testeEmptyCells() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        
        sheet.withRow().withCell("Coluna 1")
                .and().withEmptyCell()
                .and().withCell("Coluna 3")
                .and().withEmptyCells(2).withCell("Coluna 6");
        spreadsheet.toFile("/tmp/planilha_empty_cells" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testePlanilhaModelo() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");

        sheet.withRow().withGroup("Titulo").withCell("MAIOR QUE O MUNDO").withMergedColumns(9).withBackgroundColor(IndexedColors.BLACK).withFontColor(IndexedColors.WHITE);

        sheet.withRow().withGroup("SubTitulo").withCell("Subtitulo do filme").withMergedColumns(9).withBackgroundColor(IndexedColors.GREY_50_PERCENT).withFontColor(IndexedColors.WHITE);

        sheet.withRow().withGroup("Especificacoes")
                .withCell("Cls").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Descri\u00E7\u00E3o").withBackgroundColor(IndexedColors.GREY_25_PERCENT).and()
                .withCell("Verba aprovada").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT).and()
                .withCell("Verba distribuida").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT).and()
                .withCell("Saldo").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT).and()
                .withCell("Previsto").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT).and()
                .withCell("Contratado").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT).and()
                .withCell("Entregue").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT).and()
                .withCell("Real").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);

        sheet.withRow().withGroup("Item 1")
                .withCell("1").bold().and()
                .withCell("Distribui\u00E7\u00E3o 9").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").bold().and()
                .withCell(BigDecimal.valueOf(163125.92)).bold().and()
                .withCell(BigDecimal.valueOf(88276.25)).bold().and()
                .withCell(BigDecimal.valueOf(61415.44)).bold().and()
                .withCell(BigDecimal.valueOf(26860.81)).bold().and()
                .withCell(BigDecimal.valueOf(24602.56)).bold().and()
                .withCell(BigDecimal.valueOf(2258.25)).bold().and()
                .withCell(BigDecimal.valueOf(26860.81)).bold();

        sheet.withRow().withGroup("Item 2")
                .withCell("01.01").and()
                .withCell("Equipe").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(14500.00)).withDataFormat("#,##0.00").and()
                .withCell(BigDecimal.valueOf(0.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 3")
                .withCell("01.01.01").and()
                .withCell("Alimenta\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(2500.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 4")
                .withCell("01.01.02").and()
                .withCell("Comunica\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(500.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 5")
                .withCell("01.01.03").and()
                .withCell("Equipe de lan\u00E7amento").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");

        sheet.withRow().withGroup("Item 6")
                .withCell("01.01.04").and()
                .withCell("Hospedagem").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(2000.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 7")
                .withCell("01.01.05").and()
                .withCell("Transporte").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(9500.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 8")
                .withCell("01.02").and()
                .withCell("C\u00F3pia 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(2490.00)).withDataFormat("#,##0.00").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 9")
                .withCell("01.02.01").and()
                .withCell("Filme").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(1860.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 10")
                .withCell("01.02.02").and()
                .withCell("Trailer").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");

        sheet.withRow().withGroup("Item 11")
                .withCell("01.02.03").and()
                .withCell("KDM 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(630.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 12")
                .withCell("01.02.03.01").and()
                .withCell("Key Delivery Message 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 13")
                .withCell("").and()
                .withCell("KDM 2022/25 - Cinesystem Villa Romana. [226730]").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        sheet.withRow().withGroup("Item 14")
                .withCell("01.02.04").and()
                .withCell("VPF").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 15")
                .withCell("01.02.05").and()
                .withCell("Produ\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 16")
                .withCell("01.03").and()
                .withCell("Publicidade 3").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(40600.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(46000.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(35675.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(10325.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(10325.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(10325.00)).withDataFormat("#,###.00");

        sheet.withRow().withGroup("Item 17")
                .withCell("01.03.01").and()
                .withCell("Material Gr\u00E1fico 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(10600.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(675.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(325.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(325.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(325.00)).withDataFormat("#,###.00");

        sheet.withRow().withGroup("Item 18")
                .withCell("01.03.02").and()
                .withCell("M\u00EDdia 2").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(30000.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(25000.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00");

        sheet.withRow().withGroup("Item 19")
                .withCell("01.03.03").and()
                .withCell("Produ\u00E7\u00E3o Audiovisual").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell("").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 20")
                .withCell("01.04").and()
                .withCell("A\u00E7\u00E3o Promocional 4").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(70000.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(25800.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(10740.44)).and()
                .withCell(BigDecimal.valueOf(15059.56)).and()
                .withCell(BigDecimal.valueOf(14259.56)).and()
                .withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(15059.56));

        sheet.withRow().withGroup("Item 21")
                .withCell("01.04.01").and()
                .withCell("Evento 2").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(10740.44)).and()
                .withCell(BigDecimal.valueOf(14259.56)).and()
                .withCell(BigDecimal.valueOf(14259.56)).and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(14259.56));

        sheet.withRow().withGroup("Item 22")
                .withCell("01.04.02").and()
                .withCell("Imprensa 2").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(15000.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00");

        sheet.withRow().withGroup("Item 23")
                .withCell("01.04.03").and()
                .withCell("Merchandising").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(45000.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 24")
                .withCell("01.04.04").and()
                .withCell("Pesquisa").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 25")
                .withCell("01.05").and()
                .withCell("Transporte").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(2800.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(0.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 26")
                .withCell("01.05.01").and()
                .withCell("Frete").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(2800.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 27")
                .withCell("01.06").and()
                .withCell("Taxas e tributos").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(18351.67)).and()
                .withCell(BigDecimal.valueOf(0.00)).withDataFormat("#,##0.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 28")
                .withCell("01.06.01").and()
                .withCell("Encargos Sociais").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(18351.67)).and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 29")
                .withCell("01.07").and()
                .withCell("Outros 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(14384.25)).and()
                .withCell(BigDecimal.valueOf(16458.25)).and()
                .withCell(BigDecimal.valueOf(15000.00)).withDataFormat("#,###.00").and()
                .withCell(BigDecimal.valueOf(1458.25)).and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(1458.25)).and()
                .withCell(BigDecimal.valueOf(1458.25));

        sheet.withRow().withGroup("Item 30")
                .withCell("01.07.01").and()
                .withCell("Classifica\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(966.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(15000.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 31")
                .withCell("01.07.02").and()
                .withCell("Fiscaliza\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(11960.00)).withDataFormat("#,###.00").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 32")
                .withCell("01.07.03").and()
                .withCell("Honor\u00E1rios").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Item 33")
                .withCell("01.07.04").and()
                .withCell("Seguro").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").and()
                .withCell(BigDecimal.valueOf(1458.25)).and()
                .withCell("").and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(1458.25)).and()
                .withCell("").and()
                .withCell(BigDecimal.valueOf(1458.25)).and()
                .withCell(BigDecimal.valueOf(1458.25));

        sheet.withRow().withGroup("Item 34")
                .withCell("01.07.05").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("").and()
                .withCell("");

        sheet.withRow().withGroup("Total")
                .withCell("Total").bold().withMergedColumns(2).withBorder().and()
                .withCell("").withFormula("SUBTOTAL(109,C4:C37)").bold().withBorderTop().and()
                .withCell("").withFormula("SUBTOTAL(109,D4:D37)").bold().withBorderTop().and()
                .withCell("").withFormula("SUBTOTAL(109,E4:E37)").bold().withBorderTop().and()
                .withCell("").withFormula("SUBTOTAL(109,F4:F37)").bold().withBorderTop().and()
                .withCell("").withFormula("SUBTOTAL(109,G4:G37)").bold().withBorderTop().and()
                .withCell("").withFormula("SUBTOTAL(109,H4:H37)").bold().withBorderTop().and()
                .withCell("").withFormula("SUBTOTAL(109,I4:I37)").bold().withBorderTop();

        sheet.freeze(2, 3);
        sheet.withAutoSizeColumns(true);

        spreadsheet.toFile("/tmp/planilha_modelo_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }

    @Test
    @Disabled
    public void testeDiego() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Diego");
        sheet.freeze(3, 3).withAutoSizeColumns(true);

        final WMXRow rowHeader = sheet.withRow();//.title();
        rowHeader.withCell("Titulo do relatório".toUpperCase()).header().withMergedColumns(5);

        final WMXRow rowTitle = sheet.withRow();//.title();
        rowTitle.withCell("ID").title().withMergedRows(2);
        rowTitle.withCell("Cls").title().withMergedRows(2);
        rowTitle.withCell("Name").title().withMergedRows(2);
        rowTitle.withCell("Ammount").title().withMergedColumns(2);

        final WMXRow rowSubtitle = sheet.withRow();//.subtitle();
        rowSubtitle.withCell("ID").subtitle();
        rowSubtitle.withCell("Cls").subtitle();
        rowSubtitle.withCell("Name").subtitle();
        rowSubtitle.withCell("Gross").subtitle().withHorizontalAligment(HorizontalAlignment.RIGHT);
        rowSubtitle.withCell("Net").subtitle().withHorizontalAligment(HorizontalAlignment.RIGHT);

        for (int i = 1; i < 10; i++) {
            final WMXRow row = sheet.withRow();
            row.withCell(String.valueOf(i));
            row.withCell(String.valueOf(i % 2));
            row.withCell("Fulano da Silva " + i).withLink("http://gremio.net");
            row.withCell(BigDecimal.valueOf(i * 1000)).currency();
            row.withCell(BigDecimal.valueOf(i * 1000 * 0.9 * -1)).currency();
        }

        final WMXRow rowSubTotal = sheet.withRow();//.footer();.bold()
        rowSubTotal.withCell("Subtotal").withMergedColumns(3).withBorderTop().bold();
        rowSubTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUBTOTAL(109,D4:D12)").totalizer();
        rowSubTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUBTOTAL(109,E4:E12)").totalizer();

        final WMXRow rowGrandTotal = sheet.withRow();//.footer();
        rowGrandTotal.withCell("Grand total").withMergedColumns(3).withBorderTop().bold();
        rowGrandTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUM(109,D4:D12)").totalizer();
        rowGrandTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUM(109,E4:E12)").totalizer();

        spreadsheet.toFile("/tmp/planilha_diego_" + LocalDateTime.now().format(FORMATTER) + ".xlsx");
    }
}