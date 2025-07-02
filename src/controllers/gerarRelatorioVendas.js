import { gerarNomeArquivo } from '../utils.js';
import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

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
