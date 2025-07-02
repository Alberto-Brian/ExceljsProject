import ExcelJS from 'exceljs';
import path from 'path';
import { gerarNomeArquivo } from '../utils.js';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

export async function gerarPlanilhaCustomizada(req, res) {
  try {
    const { titulo, dados, colunas } = req.body;

    if (!titulo || !dados || !colunas) {
      return res.status(400).json({
        success: false,
        message: 'Campos obrigatórios: titulo, dados, colunas'
      });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(titulo);

    worksheet.columns = colunas.map(col => ({
      header: col.header,
      key: col.key,
      width: col.width || 15
    }));

    dados.forEach(item => worksheet.addRow(item));

    worksheet.getRow(1).font = { bold: true };

    const nomeArquivo = gerarNomeArquivo('planilha-personalizada');
    const caminho = path.join(__dirname, '..', '..', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      message: 'Planilha personalizada gerada com sucesso!',
      downloadUrl: `/downloads/${nomeArquivo}`
    });

  } catch (error) {
    res.status(500).json({ success: false, message: 'Erro ao gerar planilha personalizada', error: error.message });
  }
}
