import { Request, Response } from 'express';
import ExcelJS from 'exceljs';
import path from 'path';
import { gerarNomeArquivo } from '../utils';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// const __filename = fileURLToPath(import.meta.url);
const __filename = require('path').resolve();
const __dirname = dirname(__filename);

interface Funcionario {
  nome: string;
  cargo: string;
  salario: number;
}

export async function gerarPlanilhaBasica(req: Request, res: Response): Promise<void> {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Funcionários');

    worksheet.columns = [
      { header: 'Nome', key: 'nome', width: 30 },
      { header: 'Cargo', key: 'cargo', width: 25 },
      { header: 'Salário', key: 'salario', width: 15 }
    ];

    const funcionarios: Funcionario[] = [
      { nome: 'Ana Silva', cargo: 'Analista', salario: 5000 },
      { nome: 'Carlos Lima', cargo: 'Gerente', salario: 8000 },
      { nome: 'Mariana Costa', cargo: 'Assistente', salario: 3000 }
    ];

    funcionarios.forEach(f => worksheet.addRow(f));

    worksheet.getColumn('salario').numFmt = 'R$ #,##0.00';
    worksheet.getRow(1).font = { bold: true };

    const nomeArquivo = gerarNomeArquivo('planilha-basica');
    const caminho = path.join(__dirname, 'project', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      message: 'Planilha básica gerada com sucesso!',
      downloadUrl: `/downloads/${nomeArquivo}`,
      total: funcionarios.length
    });

  } catch (error) {
    res.status(500).json({ 
      success: false, 
      message: 'Erro ao gerar planilha básica', 
      error: error instanceof Error ? error.message : 'Erro desconhecido'
    });
  }
}