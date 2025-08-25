import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { Request, Response } from 'express';
import { gerarNomeArquivo } from '../utils';

// --- CORREÇÃO: Definir __filename e __dirname para Módulos ES ---
// const __filename = fileURLToPath(import.meta.url);
const __filename = require('path').resolve();
const __dirname = dirname(__filename);

// --- Interfaces TypeScript ---

interface CellStyle {
  text?: string;
  value?: string | number;
  font?: Partial<ExcelJS.Font>;
  alignment?: Partial<ExcelJS.Alignment>;
  fill?: ExcelJS.Fill;
  border?: Partial<ExcelJS.Borders>;
  numFmt?: string;
}

interface VeiculoDados {
  tarifa: number;
  especie: number;
  tpa: number;
  isento: number;
}

interface Veiculos {
  [classe: string]: VeiculoDados;
}

interface Valores {
  saldoInicial: number;
  valorDeclarado: number;
}

interface DadosRelatorio {
  chefeTurno: string;
  cabine: string;
  ref: string;
  dataAbertura: string;
  dataFechamento: string;
  operador: string;
  veiculos: Veiculos;
  valores: Valores;
}

interface TotaisTabela {
  especie: number;
  tpa: number;
  isento: number;
  valorEspecie: number;
  valorTPA: number;
  valorIsento: number;
  valorTotal: number;
  totalVeiculos: number;
}

interface AlignmentConfig {
  titleStartCol: string;
  titleEndCol: string;
  labelStartCol: string;
  labelEndCol: string;
  valueCol: string;
}

interface ResumoFinanceiro {
  totalGeral: number;
  diferenca: number;
}

interface ResponseData {
  success: boolean;
  message: string;
  downloadUrl?: string;
  dados?: {
    totalVeiculos: number;
    totalArrecadado: number;
    diferenca: number;
    operador: string;
    periodo: string;
  };
  error?: string;
}

type AlignmentOption = 'left' | 'right' | 'center';

// --- Funções de Utilitário para Excel ---

const applyCellStyles = (cell: ExcelJS.Cell, options: CellStyle): void => {
  const { text, value, font = {}, alignment = {}, fill, border, numFmt } = options;
  
  if (text !== undefined) cell.value = text;
  if (value !== undefined) cell.value = value;
  
  cell.font = { name: 'Calibri', size: 10, ...font };
  cell.alignment = { vertical: 'middle', ...alignment };
  
  if (fill) cell.fill = fill;
  if (border) cell.border = border;
  if (numFmt) cell.numFmt = numFmt;
};

const applyBordersToRange = (
  worksheet: ExcelJS.Worksheet, 
  range: string, 
  borderStyle: Partial<ExcelJS.Borders> = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  }
): void => {
  const [startCell, endCell] = range.split(':');
  const startRow = parseInt(startCell.match(/\d+/)![0], 10);
  const startCol = worksheet.getColumn(startCell.match(/[A-Z]+/)![0]).number;
  const endRow = parseInt(endCell.match(/\d+/)![0], 10);
  const endCol = worksheet.getColumn(endCell.match(/[A-Z]+/)![0]).number;
  
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const cell = worksheet.getCell(r, c);
      cell.border = { ...cell.border, ...borderStyle };
    }
  }
};

// --- Funções de Geração de Secções do Relatório ---

const criarCabecalho = (worksheet: ExcelJS.Worksheet, dados: DadosRelatorio): void => {
  worksheet.mergeCells('A1:C3');
  applyCellStyles(worksheet.getCell('A1'), {
    text: 'FROE\nFundo Rodoviário\ne Obras de Emergência',
    font: { bold: true, size: 10 },
    alignment: { horizontal: 'left', vertical: 'top', wrapText: true }
  });

  worksheet.mergeCells('D1:I2');
  applyCellStyles(worksheet.getCell('D1'), {
    text: 'Relatório de Fecho de Caixa',
    font: { size: 18, bold: true },
    alignment: { horizontal: 'center' }
  });

  const labelStyle: Partial<CellStyle> = {
    font: { bold: true, size: 10 },
    alignment: { horizontal: 'left' }
  };

  applyCellStyles(worksheet.getCell('A4'), { text: 'Chefe de Turno:', ...labelStyle });
  applyCellStyles(worksheet.getCell('B4'), { text: dados.chefeTurno });
  applyCellStyles(worksheet.getCell('A5'), { text: 'Operador(a):', ...labelStyle });
  applyCellStyles(worksheet.getCell('B5'), { text: dados.operador });
  applyCellStyles(worksheet.getCell('E4'), {
    text: dados.cabine,
    ...labelStyle,
    alignment: { horizontal: 'left' }
  });
  applyCellStyles(worksheet.getCell('E5'), {
    text: `Ref: ${dados.ref}`,
    ...labelStyle,
    alignment: { horizontal: 'left' }
  });
  applyCellStyles(worksheet.getCell('A6'), { text: 'Data de Abertura:', ...labelStyle });
  applyCellStyles(worksheet.getCell('B6'), { text: dados.dataAbertura });
  applyCellStyles(worksheet.getCell('A7'), { text: 'Data de Fechamento:', ...labelStyle });
  applyCellStyles(worksheet.getCell('B7'), { text: dados.dataFechamento });
  
  worksheet.getRow(8).height = 10;
};

const criarTabelaVeiculos = (worksheet: ExcelJS.Worksheet, veiculos: Veiculos): TotaisTabela => {
  const startRow = 9;
  
  worksheet.mergeCells(`A${startRow}:I${startRow}`);
  applyCellStyles(worksheet.getCell(`A${startRow}`), {
    text: 'REGISTO DE VEÍCULOS & VALOR PAGO',
    font: { size: 12, bold: true },
    alignment: { horizontal: 'center' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }
  });

  const headerRow1 = worksheet.getRow(startRow + 1);
  headerRow1.values = ['CLASSE', 'ESPÉCIE', null, 'TPA/RUPE', null, 'ISENTO', null, 'TOTAL', null];
  
  worksheet.mergeCells(`B${startRow + 1}:C${startRow + 1}`);
  worksheet.mergeCells(`D${startRow + 1}:E${startRow + 1}`);
  worksheet.mergeCells(`F${startRow + 1}:G${startRow + 1}`);
  worksheet.mergeCells(`H${startRow + 1}:I${startRow + 1}`);
  
  headerRow1.eachCell((cell) => applyCellStyles(cell, {
    font: { bold: true },
    alignment: { horizontal: 'center' }
  }));

  const headerRow2 = worksheet.getRow(startRow + 2);
  headerRow2.values = [
    '', 'Nº de veiculo', 'Valor (Kz)', 'Nº de veiculo', 'Valor (Kz)',
    'Nº de veiculo', 'Valor (Kz)', 'Nº de veiculo', 'Valor (Kz)'
  ];
  
  headerRow2.eachCell((cell) => applyCellStyles(cell, {
    font: { bold: true, size: 9 },
    alignment: { horizontal: 'center', wrapText: true }
  }));

  const totals: TotaisTabela = {
    especie: 0,
    tpa: 0,
    isento: 0,
    valorEspecie: 0,
    valorTPA: 0,
    valorIsento: 0,
    valorTotal: 0,
    totalVeiculos: 0
  };

  Object.entries(veiculos).forEach(([classe, dados], index) => {
    const rowNumber = startRow + 3 + index;
    const valorEspecie = dados.especie * dados.tarifa;
    const valorTPA = dados.tpa * dados.tarifa;
    const valorIsento = dados.isento * dados.tarifa;

    worksheet.getCell(`A${rowNumber}`).value = `${classe} - ${dados.tarifa}`;
    worksheet.getCell(`B${rowNumber}`).value = dados.especie || 0;
    worksheet.getCell(`C${rowNumber}`).value = valorEspecie;
    worksheet.getCell(`D${rowNumber}`).value = dados.tpa || 0;
    worksheet.getCell(`E${rowNumber}`).value = valorTPA;
    worksheet.getCell(`F${rowNumber}`).value = dados.isento || 0;
    worksheet.getCell(`G${rowNumber}`).value = valorIsento;
    worksheet.getCell(`H${rowNumber}`).value = dados.especie + dados.tpa + dados.isento;
    worksheet.getCell(`I${rowNumber}`).value = valorEspecie + valorTPA + valorIsento;

    totals.especie += dados.especie;
    totals.tpa += dados.tpa;
    totals.isento += dados.isento;
    totals.valorEspecie += valorEspecie;
    totals.valorTPA += valorTPA;
    totals.valorIsento += valorIsento;
  });

  totals.totalVeiculos = totals.especie + totals.tpa + totals.isento;
  totals.valorTotal = totals.valorEspecie + totals.valorTPA + totals.valorIsento;

  const totalRowNumber = startRow + 3 + Object.keys(veiculos).length;
  const totalRow = worksheet.getRow(totalRowNumber);
  totalRow.values = [
    'Total', totals.especie, totals.valorEspecie, totals.tpa, totals.valorTPA,
    totals.isento, totals.valorIsento, totals.totalVeiculos, totals.valorTotal
  ];
  
  totalRow.eachCell((cell) => applyCellStyles(cell, { font: { bold: true } }));

  const tableRange = `A${startRow}:I${totalRowNumber}`;
  applyBordersToRange(worksheet, tableRange);

  // Aplicar formatação às células
  for (let r = startRow + 3; r <= totalRowNumber; r++) {
    worksheet.getCell(r, 1).alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Colunas de valores (3, 5, 7, 9)
    for (const c of [3, 5, 7, 9]) {
      const cell = worksheet.getCell(r, c);
      cell.alignment = { horizontal: 'center' };
      cell.numFmt = '#,##0.00 "Kz"';
    }
    
    // Colunas de números (2, 4, 6, 8)
    for (const c of [2, 4, 6, 8]) {
      const cell = worksheet.getCell(r, c);
      cell.alignment = { horizontal: 'center' };
      cell.numFmt = '0';
    }
  }

  return totals;
};

const getAlignmentCols = (alignment: AlignmentOption = 'center'): AlignmentConfig => {
  switch (alignment) {
    case 'left':
      return {
        titleStartCol: 'A',
        titleEndCol: 'E',
        labelStartCol: 'A',
        labelEndCol: 'D',
        valueCol: 'E'
      };
    case 'right':
      return {
        titleStartCol: 'E',
        titleEndCol: 'I',
        labelStartCol: 'E',
        labelEndCol: 'H',
        valueCol: 'I'
      };
    case 'center':
    default:
      return {
        titleStartCol: 'C',
        titleEndCol: 'G',
        labelStartCol: 'C',
        labelEndCol: 'F',
        valueCol: 'G'
      };
  }
};

const criarResumoFinanceiro = (
  worksheet: ExcelJS.Worksheet,
  valores: Valores,
  totalsTabela: TotaisTabela,
  alignmentConfig: AlignmentConfig
): ResumoFinanceiro => {
  const startRow = worksheet.lastRow!.number + 2;
  const { titleStartCol, titleEndCol, labelStartCol, labelEndCol, valueCol } = alignmentConfig;

  worksheet.mergeCells(`${titleStartCol}${startRow}:${titleEndCol}${startRow}`);
  applyCellStyles(worksheet.getCell(`${titleStartCol}${startRow}`), {
    text: 'VALORES ARRECADADOS (AOA)',
    font: { size: 12, bold: true },
    alignment: { horizontal: 'center' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }
  });

  const addFinanceRow = (label: string, value: number, isBold: boolean = false): void => {
    const row = worksheet.addRow([]);
    const rowNumber = row.number;
    
    worksheet.mergeCells(`${labelStartCol}${rowNumber}:${labelEndCol}${rowNumber}`);
    const labelCell = worksheet.getCell(`${labelStartCol}${rowNumber}`);
    applyCellStyles(labelCell, {
      text: label,
      font: { bold: isBold },
      alignment: { horizontal: 'left', vertical: 'middle' }
    });

    const valueCell = worksheet.getCell(`${valueCol}${rowNumber}`);
    applyCellStyles(valueCell, {
      value: value,
      font: { bold: isBold },
      alignment: { horizontal: 'right', vertical: 'middle' },
      numFmt: '#,##0.00 "Kz"'
    });
  };

  addFinanceRow('Saldo inicial (para troco)', valores.saldoInicial);
  addFinanceRow('Total em Espécie', totalsTabela.valorEspecie);
  addFinanceRow('Total em TPA/RUPE', totalsTabela.valorTPA);
  addFinanceRow('Total em Isentos', totalsTabela.valorIsento);
  
  const totalGeral = totalsTabela.valorEspecie + totalsTabela.valorTPA;
  addFinanceRow('Total Geral (Espécie + TPA)', totalGeral, true);
  addFinanceRow('Valor Declarado', valores.valorDeclarado);
  
  const diferenca = valores.valorDeclarado - totalGeral;
  addFinanceRow('Diferença', diferenca, true);

  const endRow = worksheet.lastRow!.number;
  applyBordersToRange(worksheet, `${titleStartCol}${startRow}:${titleEndCol}${endRow}`);

  return { totalGeral, diferenca };
};

const criarRodape = (worksheet: ExcelJS.Worksheet): void => {
  let startRow = worksheet.lastRow!.number + 3;
  
  applyCellStyles(worksheet.getCell(`A${startRow}`), {
    text: 'Observação:',
    font: { bold: true }
  });
  
  startRow += 4;
  
  worksheet.mergeCells(`A${startRow}:C${startRow}`);
  applyCellStyles(worksheet.getCell(`A${startRow}`), {
    text: '_______________________',
    alignment: { horizontal: 'center' }
  });
  
  worksheet.mergeCells(`G${startRow}:I${startRow}`);
  applyCellStyles(worksheet.getCell(`G${startRow}`), {
    text: '_______________________',
    alignment: { horizontal: 'center' }
  });
  
  startRow++;
  
  worksheet.mergeCells(`A${startRow}:C${startRow}`);
  applyCellStyles(worksheet.getCell(`A${startRow}`), {
    text: 'Operador(a)',
    alignment: { horizontal: 'center' }
  });
  
  worksheet.mergeCells(`G${startRow}:I${startRow}`);
  applyCellStyles(worksheet.getCell(`G${startRow}`), {
    text: 'Chefe de Turno',
    alignment: { horizontal: 'center' }
  });
  
  startRow += 2;
  
  worksheet.mergeCells(`A${startRow}:I${startRow}`);
  applyCellStyles(worksheet.getCell(`A${startRow}`), {
    text: 'X-access portagens - Portagem da barra do Kwanza',
    font: { size: 8 },
    alignment: { horizontal: 'center' }
  });
};

// --- Controlador Principal ---

export async function gerarFechoDeCaixa(req: Request, res: Response): Promise<void> {
  try {
    const dadosMockados: DadosRelatorio = {
      chefeTurno: 'Luciano Alberto',
      cabine: 'Cabine 2 / Pista 2',
      ref: 'REF001',
      dataAbertura: '21-01-2025 00:14:46',
      dataFechamento: '21-01-2025 07:10:53',
      operador: 'Ilidio Pundi',
      veiculos: {
        A: { tarifa: 150, especie: 0, tpa: 0, isento: 0 },
        A1: { tarifa: 100, especie: 0, tpa: 0, isento: 0 },
        B: { tarifa: 500, especie: 21, tpa: 1, isento: 3 },
        C: { tarifa: 2000, especie: 38, tpa: 3, isento: 1 },
        C1: { tarifa: 1000, especie: 56, tpa: 0, isento: 0 },
      },
      valores: {
        saldoInicial: 0.00,
        valorDeclarado: 221500.00,
      },
    };

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Fecho de Caixa');

    const alignmentOption: AlignmentOption = 'right';
    const alignmentConfig = getAlignmentCols(alignmentOption);

    criarCabecalho(worksheet, dadosMockados);
    const totaisTabela = criarTabelaVeiculos(worksheet, dadosMockados.veiculos);
    const { totalGeral, diferenca } = criarResumoFinanceiro(
      worksheet,
      dadosMockados.valores,
      totaisTabela,
      alignmentConfig
    );
    criarRodape(worksheet);

    // --- CORREÇÃO: Aumentar a largura das colunas ---
    worksheet.columns = [
      { width: 20 }, // A: Classe
      { width: 10 }, // B: Nº Espécie
      { width: 18 }, // C: Valor Espécie
      { width: 10 }, // D: Nº TPA
      { width: 18 }, // E: Valor TPA
      { width: 10 }, // F: Nº Isento
      { width: 18 }, // G: Valor Isento
      { width: 10 }, // H: Nº Total
      { width: 18 }  // I: Valor Total
    ];

    const nomeArquivo = gerarNomeArquivo(`fecho-caixa-${alignmentOption}`);
    const caminho = path.join(__dirname, 'project', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    const responseData: ResponseData = {
      success: true,
      message: 'Relatório de fecho de caixa gerado com sucesso!',
      downloadUrl: `/downloads/${nomeArquivo}`,
      dados: {
        totalVeiculos: totaisTabela.totalVeiculos,
        totalArrecadado: totalGeral,
        diferenca: diferenca,
        operador: dadosMockados.operador,
        periodo: `${dadosMockados.dataAbertura} - ${dadosMockados.dataFechamento}`,
      },
    };

    res.json(responseData);
  } catch (error) {
    console.error('Erro ao gerar planilha:', error);
    
    const errorResponse: ResponseData = {
      success: false,
      message: 'Erro ao gerar planilha de fecho de caixa',
      error: error instanceof Error ? error.message : 'Erro desconhecido',
    };

    res.status(500).json(errorResponse);
  }
}