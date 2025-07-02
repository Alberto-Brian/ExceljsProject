import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Função utilitária
function gerarNomeArquivo(prefixo) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  return `${prefixo}-${timestamp}.xlsx`;
}

// 📄 Planilha Básica
export async function gerarPlanilhaBasica(req, res) {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Básico');
    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Nome', key: 'nome', width: 25 }
    ];
    worksheet.addRow({ id: 1, nome: 'Exemplo' });

    const nomeArquivo = gerarNomeArquivo('planilha-basica');
    const caminho = path.join(__dirname, '..', '..', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      message: 'Planilha básica gerada',
      downloadUrl: `/downloads/${nomeArquivo}`
    });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Erro ao gerar planilha básica', error: error.message });
  }
}

// 📊 Relatório de Vendas
export async function gerarRelatorioVendas(req, res) {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Relatório de Vendas');

    worksheet.mergeCells('A1:F1');
    worksheet.getCell('A1').value = 'RELATÓRIO DE VENDAS - ' + new Date().getFullYear();
    worksheet.getCell('A1').font = { size: 16, bold: true };
    worksheet.getCell('A1').alignment = { horizontal: 'center' };
    worksheet.addRow([]);

    const headerRow = worksheet.addRow(['Mês', 'Vendas', 'Meta', 'Diferença', '% Meta', 'Status']);

    const vendas = [
      { mes: 'Janeiro', vendas: 150000, meta: 120000 },
      { mes: 'Fevereiro', vendas: 180000, meta: 150000 },
      { mes: 'Março', vendas: 220000, meta: 180000 },
      { mes: 'Abril', vendas: 195000, meta: 200000 },
      { mes: 'Maio', vendas: 280000, meta: 250000 },
      { mes: 'Junho', vendas: 320000, meta: 280000 }
    ];

    vendas.forEach((item, index) => {
      const linha = index + 4;
      worksheet.addRow([
        item.mes,
        item.vendas,
        item.meta,
        { formula: `B${linha}-C${linha}` },
        { formula: `B${linha}/C${linha}*100` },
        { formula: `IF(B${linha}>=C${linha},"Meta Atingida","Abaixo da Meta")` }
      ]);
    });

    const totalLinha = vendas.length + 4;
    worksheet.addRow([]);
    worksheet.addRow([
      'TOTAL',
      { formula: `SUM(B4:B${totalLinha - 1})` },
      { formula: `SUM(C4:C${totalLinha - 1})` },
      { formula: `SUM(D4:D${totalLinha - 1})` },
      { formula: `AVERAGE(E4:E${totalLinha - 1})` },
      ''
    ]);

    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };

    worksheet.getColumn(2).numFmt = 'R$ #,##0';
    worksheet.getColumn(3).numFmt = 'R$ #,##0';
    worksheet.getColumn(4).numFmt = 'R$ #,##0';
    worksheet.getColumn(5).numFmt = '0.0%';
    worksheet.columns.forEach(col => col.width = 15);

    const nomeArquivo = gerarNomeArquivo('relatorio-vendas');
    const caminho = path.join(__dirname, '..', '..', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      message: 'Relatório gerado!',
      downloadUrl: `/downloads/${nomeArquivo}`,
      totalVendas: vendas.reduce((s, i) => s + i.vendas, 0)
    });

  } catch (error) {
    res.status(500).json({ success: false, message: 'Erro ao gerar relatório de vendas', error: error.message });
  }
}

// 📦 Lista de Produtos
export async function gerarListaProdutos(req, res) {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Produtos');

    worksheet.columns = [
      { header: 'Código', key: 'codigo', width: 12 },
      { header: 'Produto', key: 'produto', width: 25 },
      { header: 'Categoria', key: 'categoria', width: 15 },
      { header: 'Preço', key: 'preco', width: 12 },
      { header: 'Estoque', key: 'estoque', width: 10 },
      { header: 'Valor Total', key: 'total', width: 15 }
    ];

    const produtos = [
      { codigo: 'P001', produto: 'Notebook', categoria: 'Informática', preco: 2500, estoque: 10 },
      { codigo: 'P002', produto: 'Mouse', categoria: 'Informática', preco: 80, estoque: 50 }
    ];

    produtos.forEach((produto, i) => {
      const linha = i + 2;
      worksheet.addRow({
        ...produto,
        total: { formula: `D${linha}*E${linha}` }
      });
    });

    worksheet.getRow(1).font = { bold: true };
    worksheet.getColumn('preco').numFmt = 'R$ #,##0.00';
    worksheet.getColumn('total').numFmt = 'R$ #,##0.00';

    const nomeArquivo = gerarNomeArquivo('lista-produtos');
    const caminho = path.join(__dirname, '..', '..', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      downloadUrl: `/downloads/${nomeArquivo}`,
      totalProdutos: produtos.length
    });

  } catch (error) {
    res.status(500).json({ success: false, message: 'Erro ao gerar lista de produtos', error: error.message });
  }
}

// 🛠️ Planilha Personalizada
export async function gerarPlanilhaCustomizada(req, res) {
  try {
    const { titulo, dados, colunas } = req.body;

    if (!titulo || !dados || !colunas) {
      return res.status(400).json({ success: false, message: 'Faltam parâmetros obrigatórios' });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(titulo);

    worksheet.columns = colunas.map(col => ({
      header: col.header,
      key: col.key,
      width: col.width || 15
    }));

    dados.forEach(d => worksheet.addRow(d));
    worksheet.getRow(1).font = { bold: true };

    const nomeArquivo = gerarNomeArquivo('planilha-personalizada');
    const caminho = path.join(__dirname, '..', '..', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      downloadUrl: `/downloads/${nomeArquivo}`,
      registros: dados.length
    });

  } catch (error) {
    res.status(500).json({ success: false, message: 'Erro ao gerar planilha personalizada', error: error.message });
  }
}
