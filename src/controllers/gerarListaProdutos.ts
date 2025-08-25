import { Request, Response } from 'express';
import ExcelJS from 'exceljs';
import path from 'path';
import { gerarNomeArquivo } from '../utils';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// const __filename = fileURLToPath(import.meta.url);
const __filename = require('path').resolve();
const __dirname = dirname(__filename);

interface Produto {
  codigo: string;
  produto: string;
  categoria: string;
  preco: number;
  estoque: number;
}

export async function gerarListaProdutos(req: Request, res: Response): Promise<void> {
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

    const produtos: Produto[] = [
      { codigo: 'P001', produto: 'Notebook', categoria: 'Informática', preco: 3000, estoque: 10 },
      { codigo: 'P002', produto: 'Mouse', categoria: 'Acessórios', preco: 100, estoque: 50 }
    ];

    produtos.forEach((p, i) => {
      const linha = i + 2;
      worksheet.addRow({
        ...p,
        total: { formula: `D${linha}*E${linha}` }
      });
    });

    worksheet.getColumn('preco').numFmt = 'R$ #,##0.00';
    worksheet.getColumn('total').numFmt = 'R$ #,##0.00';
    worksheet.getRow(1).font = { bold: true };

    const nomeArquivo = gerarNomeArquivo('lista-produtos');
    const caminho = path.join(__dirname, 'project', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      message: 'Lista de produtos gerada com sucesso!',
      downloadUrl: `/downloads/${nomeArquivo}`
    });

  } catch (error) {
    res.status(500).json({ 
      success: false, 
      message: 'Erro ao gerar lista de produtos', 
      error: error instanceof Error ? error.message : 'Erro desconhecido'
    });
  }
}