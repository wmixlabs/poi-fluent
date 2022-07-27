package br.com.wmixvideo.poi;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
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
        row.withCell("Esse e um merge de 3 celulas").withMergedColumns(3);
        row.withCell("Esse e um merge de 2 celulas e 2 linhas").withMergedRows(2).withMergedColumns(2);
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
        dfRowIV.withCell("Linha 4 agrupada").withBackgroundColor(IndexedColors.GREY_50_PERCENT);
        dfRowIV.withCell("Celula 1").withBackgroundColor(IndexedColors.GREY_50_PERCENT);
        dfRowIV.withCell("Celula 2").withBackgroundColor(IndexedColors.GREY_50_PERCENT);
        dfRowIV.withCell("Celula 3").withBackgroundColor(IndexedColors.GREY_50_PERCENT);
        dfRowIV.withCell("Celula 4").withBackgroundColor(IndexedColors.GREY_50_PERCENT);
        dfRowIV.withCell("Celula 5").withBackgroundColor(IndexedColors.GREY_50_PERCENT);

        final WMXRow dfRowV = sheet.withRow().withGroup("Agrupador2");
        dfRowV.withCell("Linha 5 agrupada");
        dfRowV.withCell("Celula 1");
        dfRowV.withCell("Celula 2");
        dfRowV.withCell("Celula 3");
        dfRowV.withCell("Celula 4");
        dfRowV.withCell("Celula 5");

        final WMXRow dfRowVI = sheet.withRow();
        dfRowVI.withCell("Linha 6 desagrupada");
        dfRowVI.withCell("Celula 1").withBackgroundColor(IndexedColors.BLUE);
        dfRowVI.withCell("Celula 2").withBackgroundColor(IndexedColors.LIGHT_BLUE);
        dfRowVI.withCell("Celula 3").withBackgroundColor(IndexedColors.BLUE1);
        dfRowVI.withCell("Celula 4").withBackgroundColor(IndexedColors.CORNFLOWER_BLUE);
        dfRowVI.withCell("Celula 5").withBackgroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE);
        dfRowVI.withCell("Celula 5").withBackgroundColor(IndexedColors.PALE_BLUE);
        dfRowVI.withCell("Celula 5").withBackgroundColor(IndexedColors.ROYAL_BLUE);
        dfRowVI.withCell("Celula 5").withBackgroundColor(IndexedColors.SKY_BLUE);

        spreadsheet.toFile("/tmp/planilha_agrupamento_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testePlanilhaModelo() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();
        final WMXSheet sheet = spreadsheet.withSheet("Teste");
        final WMXRow rowTitulo = sheet.withRow().withGroup("Titulo");
        rowTitulo.withCell("MAIOR QUE O MUNDO").withMergedColumns(9).withBackgroundColor(IndexedColors.BLACK).withFontColor(IndexedColors.WHITE);

        final WMXRow rowSubTitulo = sheet.withRow().withGroup("SubTitulo");
        rowSubTitulo.withCell("Subtitulo do filme").withMergedColumns(9).withBackgroundColor(IndexedColors.GREY_50_PERCENT).withFontColor(IndexedColors.WHITE);

        final WMXRow especificacoes = sheet.withRow().withGroup("Especificacoes");
        especificacoes.withCell("Cls").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        especificacoes.withCell("Descri\u00E7\u00E3o").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        especificacoes.withCell("Verba aprovada").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);
        especificacoes.withCell("Verba distribuida").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);
        especificacoes.withCell("Saldo").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);
        especificacoes.withCell("Previsto").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);
        especificacoes.withCell("Contratado").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);
        especificacoes.withCell("Entregue").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);
        especificacoes.withCell("Real").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);

        final WMXRow dfRowI = sheet.withRow().withGroup("Item 1");
        dfRowI.withCell("1").bold();
        dfRowI.withCell("Distribui\u00E7\u00E3o 9").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#").bold();
        dfRowI.withCell(BigDecimal.valueOf(163125.92)).bold();
        dfRowI.withCell(BigDecimal.valueOf(88276.25)).bold();
        dfRowI.withCell(BigDecimal.valueOf(61415.44)).bold();
        dfRowI.withCell(BigDecimal.valueOf(26860.81)).bold();
        dfRowI.withCell(BigDecimal.valueOf(24602.56)).bold();
        dfRowI.withCell(BigDecimal.valueOf(2258.25)).bold();
        dfRowI.withCell(BigDecimal.valueOf(26860.81)).bold();

        final WMXRow dfRowII = sheet.withRow().withGroup("Item 2");
        dfRowII.withCell("01.01");
        dfRowII.withCell("Equipe").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowII.withCell(BigDecimal.valueOf(14500.00)).withDataFormat("#,##0.00");
        dfRowII.withCell(BigDecimal.valueOf(0.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowIII = sheet.withRow().withGroup("Item 3");
        dfRowIII.withCell("01.01.01");
        dfRowIII.withCell("Alimenta\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowIII.withCell(BigDecimal.valueOf(2500.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowIV = sheet.withRow().withGroup("Item 4");
        dfRowIV.withCell("01.01.02");
        dfRowIV.withCell("Comunica\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowIV.withCell(BigDecimal.valueOf(500.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowV = sheet.withRow().withGroup("Item 5");
        dfRowV.withCell("01.01.03");
        dfRowV.withCell("Equipe de lan\u00E7amento").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");

        final WMXRow dfRowVI = sheet.withRow().withGroup("Item 6");
        dfRowVI.withCell("01.01.04");
        dfRowVI.withCell("Hospedagem").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowVI.withCell(BigDecimal.valueOf(2000.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowVII = sheet.withRow().withGroup("Item 7");
        dfRowVII.withCell("01.01.05");
        dfRowVII.withCell("Transporte").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowVII.withCell(BigDecimal.valueOf(9500.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowVIII = sheet.withRow().withGroup("Item 8");
        dfRowVIII.withCell("01.02");
        dfRowVIII.withCell("C\u00F3pia 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowVIII.withCell(BigDecimal.valueOf(2490.00)).withDataFormat("#,##0.00");
        dfRowVIII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowVIII.withCell("");
        dfRowVIII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowVIII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowVIII.withCell("");
        dfRowVIII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowIX = sheet.withRow().withGroup("Item 9");
        dfRowIX.withCell("01.02.01");
        dfRowIX.withCell("Filme").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowIX.withCell(BigDecimal.valueOf(1860.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowX = sheet.withRow().withGroup("Item 10");
        dfRowX.withCell("01.02.02");
        dfRowX.withCell("Trailer").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");

        final WMXRow dfRowXI = sheet.withRow().withGroup("Item 11");
        dfRowXI.withCell("01.02.03");
        dfRowXI.withCell("KDM 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXI.withCell(BigDecimal.valueOf(630.00)).withDataFormat("#,##0.00");
        dfRowXI.withCell("");
        dfRowXI.withCell("");
        dfRowXI.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowXI.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowXI.withCell("");
        dfRowXI.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowXII = sheet.withRow().withGroup("Item 12");
        dfRowXII.withCell("01.02.03.01");
        dfRowXII.withCell("Key Delivery Message 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXII.withCell("");
        dfRowXII.withCell("");
        dfRowXII.withCell("");
        dfRowXII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowXII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowXII.withCell("");
        dfRowXII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowXIII = sheet.withRow().withGroup("Item 13");
        dfRowXIII.withCell("");
        dfRowXIII.withCell("KDM 2022/25 - Cinesystem Villa Romana. [226730]");
        dfRowXIII.withCell("");
        dfRowXIII.withCell("");
        dfRowXIII.withCell("");
        dfRowXIII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowXIII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");
        dfRowXIII.withCell("");
        dfRowXIII.withCell(BigDecimal.valueOf(18.00)).withDataFormat("#,##0.00");

        final WMXRow dfRowXIV = sheet.withRow().withGroup("Item 14");
        dfRowXIV.withCell("01.02.04");
        dfRowXIV.withCell("VPF").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXIV.withCell("");
        dfRowXIV.withCell("");
        dfRowXIV.withCell("");
        dfRowXIV.withCell("");
        dfRowXIV.withCell("");
        dfRowXIV.withCell("");
        dfRowXIV.withCell("");

        final WMXRow dfRowXV = sheet.withRow().withGroup("Item 15");
        dfRowXV.withCell("01.02.05");
        dfRowXV.withCell("Produ\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXV.withCell("");
        dfRowXV.withCell("");
        dfRowXV.withCell("");
        dfRowXV.withCell("");
        dfRowXV.withCell("");
        dfRowXV.withCell("");
        dfRowXV.withCell("");

        final WMXRow dfRowXVI = sheet.withRow().withGroup("Item 16");
        dfRowXVI.withCell("01.03");
        dfRowXVI.withCell("Publicidade 3").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXVI.withCell(BigDecimal.valueOf(40600.00)).withDataFormat("#,###.00");
        dfRowXVI.withCell(BigDecimal.valueOf(46000.00)).withDataFormat("#,###.00");
        dfRowXVI.withCell(BigDecimal.valueOf(35675.00)).withDataFormat("#,###.00");
        dfRowXVI.withCell(BigDecimal.valueOf(10325.00)).withDataFormat("#,###.00");
        dfRowXVI.withCell(BigDecimal.valueOf(10325.00)).withDataFormat("#,###.00");
        dfRowXVI.withCell("");
        dfRowXVI.withCell(BigDecimal.valueOf(10325.00)).withDataFormat("#,###.00");

        final WMXRow dfRowXVII = sheet.withRow().withGroup("Item 17");
        dfRowXVII.withCell("01.03.01");
        dfRowXVII.withCell("Material Gr\u00E1fico 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXVII.withCell(BigDecimal.valueOf(10600.00)).withDataFormat("#,###.00");
        dfRowXVII.withCell("");
        dfRowXVII.withCell(BigDecimal.valueOf(675.00)).withDataFormat("#,###.00");
        dfRowXVII.withCell(BigDecimal.valueOf(325.00)).withDataFormat("#,###.00");
        dfRowXVII.withCell(BigDecimal.valueOf(325.00)).withDataFormat("#,###.00");
        dfRowXVII.withCell("");
        dfRowXVII.withCell(BigDecimal.valueOf(325.00)).withDataFormat("#,###.00");

        final WMXRow dfRowXVIII = sheet.withRow().withGroup("Item 18");
        dfRowXVIII.withCell("01.03.02");
        dfRowXVIII.withCell("M\u00EDdia 2").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXVIII.withCell(BigDecimal.valueOf(30000.00)).withDataFormat("#,###.00");
        dfRowXVIII.withCell("");
        dfRowXVIII.withCell(BigDecimal.valueOf(25000.00)).withDataFormat("#,###.00");
        dfRowXVIII.withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00");
        dfRowXVIII.withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00");
        dfRowXVIII.withCell("");
        dfRowXVIII.withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00");

        final WMXRow dfRowXIX = sheet.withRow().withGroup("Item 19");
        dfRowXIX.withCell("01.03.03");
        dfRowXIX.withCell("Produ\u00E7\u00E3o Audiovisual").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXIX.withCell("");
        dfRowXIX.withCell("");
        dfRowXIX.withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00");
        dfRowXIX.withCell("");
        dfRowXIX.withCell("");
        dfRowXIX.withCell("");
        dfRowXIX.withCell("");

        final WMXRow dfRowXX = sheet.withRow().withGroup("Item 20");
        dfRowXX.withCell("01.04");
        dfRowXX.withCell("A\u00E7\u00E3o Promocional 4").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXX.withCell(BigDecimal.valueOf(70000.00)).withDataFormat("#,###.00");
        dfRowXX.withCell(BigDecimal.valueOf(25800.00)).withDataFormat("#,###.00");
        dfRowXX.withCell(BigDecimal.valueOf(10740.44));
        dfRowXX.withCell(BigDecimal.valueOf(15059.56));
        dfRowXX.withCell(BigDecimal.valueOf(14259.56));
        dfRowXX.withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00");
        dfRowXX.withCell(BigDecimal.valueOf(15059.56));

        final WMXRow dfRowXXI = sheet.withRow().withGroup("Item 21");
        dfRowXXI.withCell("01.04.01");
        dfRowXXI.withCell("Evento 2").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXI.withCell(BigDecimal.valueOf(10000.00)).withDataFormat("#,###.00");
        dfRowXXI.withCell("");
        dfRowXXI.withCell(BigDecimal.valueOf(10740.44));
        dfRowXXI.withCell(BigDecimal.valueOf(14259.56));
        dfRowXXI.withCell(BigDecimal.valueOf(14259.56));
        dfRowXXI.withCell("");
        dfRowXXI.withCell(BigDecimal.valueOf(14259.56));

        final WMXRow dfRowXXII = sheet.withRow().withGroup("Item 22");
        dfRowXXII.withCell("01.04.02");
        dfRowXXII.withCell("Imprensa 2").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXII.withCell(BigDecimal.valueOf(15000.00)).withDataFormat("#,###.00");
        dfRowXXII.withCell("");
        dfRowXXII.withCell("");
        dfRowXXII.withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00");
        dfRowXXII.withCell("");
        dfRowXXII.withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00");
        dfRowXXII.withCell(BigDecimal.valueOf(800.00)).withDataFormat("#,###.00");

        final WMXRow dfRowXXIII = sheet.withRow().withGroup("Item 23");
        dfRowXXIII.withCell("01.04.03");
        dfRowXXIII.withCell("Merchandising").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXIII.withCell(BigDecimal.valueOf(45000.00)).withDataFormat("#,###.00");
        dfRowXXIII.withCell("");
        dfRowXXIII.withCell("");
        dfRowXXIII.withCell("");
        dfRowXXIII.withCell("");
        dfRowXXIII.withCell("");
        dfRowXXIII.withCell("");

        final WMXRow dfRowXXIV = sheet.withRow().withGroup("Item 24");
        dfRowXXIV.withCell("01.04.04");
        dfRowXXIV.withCell("Pesquisa").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXIV.withCell("");
        dfRowXXIV.withCell("");
        dfRowXXIV.withCell("");
        dfRowXXIV.withCell("");
        dfRowXXIV.withCell("");
        dfRowXXIV.withCell("");
        dfRowXXIV.withCell("");

        final WMXRow dfRowXXV = sheet.withRow().withGroup("Item 25");
        dfRowXXV.withCell("01.05");
        dfRowXXV.withCell("Transporte").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXV.withCell(BigDecimal.valueOf(2800.00)).withDataFormat("#,###.00");
        dfRowXXV.withCell(BigDecimal.valueOf(0.00)).withDataFormat("#,##0.00");
        dfRowXXV.withCell("");
        dfRowXXV.withCell("");
        dfRowXXV.withCell("");
        dfRowXXV.withCell("");
        dfRowXXV.withCell("");

        final WMXRow dfRowXXVI = sheet.withRow().withGroup("Item 26");
        dfRowXXVI.withCell("01.05.01");
        dfRowXXVI.withCell("Frete").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXVI.withCell(BigDecimal.valueOf(2800.00)).withDataFormat("#,###.00");
        dfRowXXVI.withCell("");
        dfRowXXVI.withCell("");
        dfRowXXVI.withCell("");
        dfRowXXVI.withCell("");
        dfRowXXVI.withCell("");
        dfRowXXVI.withCell("");

        final WMXRow dfRowXXVII = sheet.withRow().withGroup("Item 27");
        dfRowXXVII.withCell("01.06");
        dfRowXXVII.withCell("Taxas e tributos").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXVII.withCell(BigDecimal.valueOf(18351.67));
        dfRowXXVII.withCell(BigDecimal.valueOf(0.00)).withDataFormat("#,##0.00");
        dfRowXXVII.withCell("");
        dfRowXXVII.withCell("");
        dfRowXXVII.withCell("");
        dfRowXXVII.withCell("");
        dfRowXXVII.withCell("");
        dfRowXXVII.withCell("");

        final WMXRow dfRowXXVIII = sheet.withRow().withGroup("Item 28");
        dfRowXXVIII.withCell("01.06.01");
        dfRowXXVIII.withCell("Encargos Sociais").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXVIII.withCell(BigDecimal.valueOf(18351.67));
        dfRowXXVIII.withCell("");
        dfRowXXVIII.withCell("");
        dfRowXXVIII.withCell("");
        dfRowXXVIII.withCell("");
        dfRowXXVIII.withCell("");
        dfRowXXVIII.withCell("");
        dfRowXXVIII.withCell("");

        final WMXRow dfRowXXIX = sheet.withRow().withGroup("Item 29");
        dfRowXXIX.withCell("01.07");
        dfRowXXIX.withCell("Outros 1").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXIX.withCell(BigDecimal.valueOf(14384.25));
        dfRowXXIX.withCell(BigDecimal.valueOf(16458.25));
        dfRowXXIX.withCell(BigDecimal.valueOf(15000.00)).withDataFormat("#,###.00");
        dfRowXXIX.withCell(BigDecimal.valueOf(1458.25));
        dfRowXXIX.withCell("");
        dfRowXXIX.withCell(BigDecimal.valueOf(1458.25));
        dfRowXXIX.withCell(BigDecimal.valueOf(1458.25));

        final WMXRow dfRowXXX = sheet.withRow().withGroup("Item 30");
        dfRowXXX.withCell("01.07.01");
        dfRowXXX.withCell("Classifica\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXX.withCell(BigDecimal.valueOf(966.00)).withDataFormat("#,###.00");
        dfRowXXX.withCell("");
        dfRowXXX.withCell(BigDecimal.valueOf(15000.00)).withDataFormat("#,###.00");
        dfRowXXX.withCell("");
        dfRowXXX.withCell("");
        dfRowXXX.withCell("");
        dfRowXXX.withCell("");

        final WMXRow dfRowXXXI = sheet.withRow().withGroup("Item 31");
        dfRowXXXI.withCell("01.07.02");
        dfRowXXXI.withCell("Fiscaliza\u00E7\u00E3o").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXXI.withCell(BigDecimal.valueOf(11960.00)).withDataFormat("#,###.00");
        dfRowXXXI.withCell("");
        dfRowXXXI.withCell("");
        dfRowXXXI.withCell("");
        dfRowXXXI.withCell("");
        dfRowXXXI.withCell("");
        dfRowXXXI.withCell("");

        final WMXRow dfRowXXXII = sheet.withRow().withGroup("Item 32");
        dfRowXXXII.withCell("01.07.03");
        dfRowXXXII.withCell("Honor\u00E1rios").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXXII.withCell("");
        dfRowXXXII.withCell("");
        dfRowXXXII.withCell("");
        dfRowXXXII.withCell("");
        dfRowXXXII.withCell("");
        dfRowXXXII.withCell("");
        dfRowXXXII.withCell("");

        final WMXRow dfRowXXXIII = sheet.withRow().withGroup("Item 33");
        dfRowXXXIII.withCell("01.07.04");
        dfRowXXXIII.withCell("Seguro").withLink("https://orcamento.wmixvideo.com.br/budget/1919/1/0#");
        dfRowXXXIII.withCell(BigDecimal.valueOf(1458.25));
        dfRowXXXIII.withCell("");
        dfRowXXXIII.withCell("");
        dfRowXXXIII.withCell(BigDecimal.valueOf(1458.25));
        dfRowXXXIII.withCell("");
        dfRowXXXIII.withCell(BigDecimal.valueOf(1458.25));
        dfRowXXXIII.withCell(BigDecimal.valueOf(1458.25));

        final WMXRow dfRowXXXIV = sheet.withRow().withGroup("Item 34");
        dfRowXXXIV.withCell("01.07.05");
        dfRowXXXIV.withCell("");
        dfRowXXXIV.withCell("");
        dfRowXXXIV.withCell("");
        dfRowXXXIV.withCell("");
        dfRowXXXIV.withCell("");
        dfRowXXXIV.withCell("");
        dfRowXXXIV.withCell("");

        final WMXRow total = sheet.withRow().withGroup("Total");
        total.withCell("Total").bold().withMergedColumns(2).withBorder();
        total.withCell("").withFormula("SUBTOTAL(109,C4:C37)").bold().withBorderTop();
        total.withCell("").withFormula("SUBTOTAL(109,D4:D37)").bold().withBorderTop();
        total.withCell("").withFormula("SUBTOTAL(109,E4:E37)").bold().withBorderTop();
        total.withCell("").withFormula("SUBTOTAL(109,F4:F37)").bold().withBorderTop();
        total.withCell("").withFormula("SUBTOTAL(109,G4:G37)").bold().withBorderTop();
        total.withCell("").withFormula("SUBTOTAL(109,H4:H37)").bold().withBorderTop();
        total.withCell("").withFormula("SUBTOTAL(109,I4:I37)").bold().withBorderTop();

        sheet.freeze(2, 3);
        sheet.withAutoSizeColumns(true);

        spreadsheet.toFile("/tmp/planilha_modelo_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }

    @Test
    @Disabled
    public void testeDiego() throws Exception {
        final WMXSpreadsheet spreadsheet = new WMXSpreadsheet();

        final WMXSheet sheet = spreadsheet.withSheet("Diego");
        sheet.freeze(3, 3).withAutoSizeColumns(true);

        final WMXRow rowHeader = sheet.withRow();//.title();
        rowHeader.withCell("Titulo do relatório".toUpperCase())
                .withHorizontalAligment(HorizontalAlignment.CENTER)
                .withBackgroundColor(IndexedColors.GREY_80_PERCENT)
                .withFontColor(IndexedColors.WHITE)
                .withMergedColumns(5)
                .bold();

        final WMXRow rowTitle = sheet.withRow();//.title();
        rowTitle.withCell("ID").withMergedRows(2).withBackgroundColor(IndexedColors.GREY_50_PERCENT).withFontColor(IndexedColors.WHITE).bold();
        rowTitle.withCell("Cls").withMergedRows(2).withBackgroundColor(IndexedColors.GREY_50_PERCENT).withFontColor(IndexedColors.WHITE).bold();
        rowTitle.withCell("Name").withMergedRows(2).withBackgroundColor(IndexedColors.GREY_50_PERCENT).withFontColor(IndexedColors.WHITE).bold();
        rowTitle.withCell("Ammount").withMergedColumns(2).withBackgroundColor(IndexedColors.GREY_50_PERCENT).withFontColor(IndexedColors.WHITE).bold();

        final WMXRow rowSubtitle = sheet.withRow();//.subtitle();
        rowSubtitle.withCell("ID").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        rowSubtitle.withCell("Cls").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        rowSubtitle.withCell("Name").withBackgroundColor(IndexedColors.GREY_25_PERCENT);
        rowSubtitle.withCell("Gross").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);
        rowSubtitle.withCell("Net").withBackgroundColor(IndexedColors.GREY_25_PERCENT).withHorizontalAligment(HorizontalAlignment.RIGHT);

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
        rowSubTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUBTOTAL(109,D4:D12)").withBorderTop().bold();
        rowSubTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUBTOTAL(109,E4:E12)").withBorderTop().bold();

        final WMXRow rowGrandTotal = sheet.withRow();//.footer();
        rowGrandTotal.withCell("Grand total").withMergedColumns(3).withBorderTop().bold();
        rowGrandTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUM(109,D4:D12)").withBorderTop().bold();
        rowGrandTotal.withCell(BigDecimal.ZERO).currency().withFormula("SUM(109,E4:E12)").withBorderTop().bold();

        spreadsheet.toFile("/tmp/planilha_diego_" + LocalDateTime.now().format(FORMATTER) + ".xls");
    }
}