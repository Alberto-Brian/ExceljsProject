// src/controllers/gerarFechoDeCaixa.js
import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { gerarNomeArquivo } from '../utils.js';

// --- Configuração de Caminhos ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// --- Funções de Utilitário para Excel ---

/**
 * Aplica um estilo de preenchimento a um conjunto de células.
 * @param {ExcelJS.Worksheet} worksheet - A folha de cálculo.
 * @param {string[]} cells - Array de referências de células (ex: ['A1', 'B2']).
 * @param {string} color - Cor de preenchimento em formato ARGB (ex: 'FFE6E6E6').
 */
const applyFill = (worksheet, cells, color) => {
  cells.forEach(cellRef => {
    worksheet.getCell(cellRef).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: color },
    };
  });
};

/**
 * Aplica bordas a um intervalo de células.
 * @param {ExcelJS.Worksheet} worksheet - A folha de cálculo.
 * @param {string} range - O intervalo de células (ex: 'A1:H10').
 */
const applyBordersToRange = (worksheet, range) => {
  const [startCell, endCell] = range.split(':');
  const startRow = parseInt(startCell.match(/\d+/)[0], 10);
  const startCol = worksheet.getColumn(startCell.match(/[A-Z]+/)[0]).number;
  const endRow = parseInt(endCell.match(/\d+/)[0], 10);
  const endCol = worksheet.getColumn(endCell.match(/[A-Z]+/)[0]).number;

  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      worksheet.getCell(row, col).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      worksheet.getCell(row, col).alignment = { horizontal: 'center', vertical: 'middle' };
    }
  }
};

// --- Funções de Geração de Secções do Relatório ---

/**
 * Cria o cabeçalho principal do relatório.
 * @param {ExcelJS.Worksheet} worksheet - A folha de cálculo.
 * @param {object} dados - Os dados do relatório.
 */
const criarCabecalho = (worksheet, dados) => {
  worksheet.mergeCells('A1:I1');
  const titleCell = worksheet.getCell('A1');
  titleCell.value = 'Relatório de Fecho de Caixa';
  titleCell.font = { size: 16, bold: true };
  titleCell.alignment = { horizontal: 'center' };

  worksheet.addRow([]); // Espaçamento
    addLabelValueRow(worksheet, 'Chefe de Turno:', dados.chefeTurno);
    addLabelValueRow(worksheet, dados.cabine, `Ref: ${dados.ref}`);
    addLabelValueRow(worksheet, 'Operador(a):', dados.operador);
    addLabelValueRow(worksheet, 'Data de Abertura:', dados.dataAbertura);
    addLabelValueRow(worksheet, 'Data de Fechamento:', dados.dataFechamento)
  worksheet.addRow([]); // Espaçamento
};

function addLabelValueRow(worksheet, label, value) {
  const row = worksheet.addRow([label, value]);
  row.getCell(1).font = { bold: true };
  return row;
}


/**
 * Cria a tabela de registo de veículos.
 * @param {ExcelJS.Worksheet} worksheet - A folha de cálculo.
 * @param {object} veiculos - Os dados dos veículos.
 * @returns {object} - Totais calculados na tabela.
 */
const criarTabelaVeiculos = (worksheet, veiculos) => {
  worksheet.mergeCells('A8:I8');
  const tableTitleCell = worksheet.getCell('A8');
  tableTitleCell.value = 'REGISTO DE VEÍCULOS & VALOR PAGO';
  tableTitleCell.font = { size: 12, bold: true };
  tableTitleCell.alignment = { horizontal: 'center' };

  // Cabeçalhos da tabela
  const headerRow = worksheet.addRow(['CLASSE', 'ESPÉCIE', '', 'TPA/RUPE', '', 'ISENTO', '', 'TOTAL', '']);
  headerRow.font = { bold: true };
  headerRow.alignment = { horizontal: 'center' };
  worksheet.mergeCells('B9:C9');
  worksheet.mergeCells('D9:E9');
  worksheet.mergeCells('F9:G9');
  worksheet.mergeCells('H9:I9');

  const subHeaderRow = worksheet.addRow(['', 'Nº', 'Valor (Kz)', 'Nº', 'Valor (Kz)', 'Nº', 'Valor (Kz)', 'Nº', 'Valor (Kz)']);
  subHeaderRow.font = { bold: true };
  subHeaderRow.alignment = { horizontal: 'center' };

  let totalGeralVeiculos = 0;
  let totalGeralValor = 0;
  let totalEspecieVeiculos = 0;
  let totalTpaVeiculos = 0;
  let totalIsentoVeiculos = 0;

  Object.entries(veiculos).forEach(([classe, dados]) => {
    const valorEspecie = dados.especie * dados.tarifa;
    const valorTPA = dados.tpa * dados.tarifa;
    const valorIsento = dados.isento * dados.tarifa;
    const totalClasseVeiculos = dados.especie + dados.tpa + dados.isento;
    const totalClasseValor = valorEspecie + valorTPA + valorIsento;

    worksheet.addRow([
      `${classe} (${dados.tarifa.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' })})`,
      dados.especie || '--',
      valorEspecie.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
      dados.tpa || '--',
      valorTPA.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
      dados.isento || '--',
      valorIsento.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
      totalClasseVeiculos,
      totalClasseValor.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
    ]);

    totalGeralVeiculos += totalClasseVeiculos;
    totalGeralValor += totalClasseValor;
    totalEspecieVeiculos += dados.especie;
    totalTpaVeiculos += dados.tpa;
    totalIsentoVeiculos += dados.isento;
  });

  return { totalGeralVeiculos, totalGeralValor, totalEspecieVeiculos, totalTpaVeiculos, totalIsentoVeiculos };
};

/**
 * Cria a secção de resumo financeiro.
 * @param {ExcelJS.Worksheet} worksheet - A folha de cálculo.
 * @param {object} valores - Os dados financeiros.
 * @returns {object} - Valores calculados como total geral e diferença.
 */
const criarResumoFinanceiro = (worksheet, valores) => {
  worksheet.addRow([]); // Espaçamento
  
  // Calcula posição centralizada (colunas C a F para usar 4 colunas)
  const startCol = 'C';
  const endCol = 'F';
  const startRow = worksheet.lastRow.number + 1;

  // Título da seção centrado
  worksheet.mergeCells(`${startCol}${startRow}:${endCol}${startRow}`);
  const summaryTitleCell = worksheet.getCell(`${startCol}${startRow}`);
  summaryTitleCell.value = 'VALORES ARRECADADOS (AOA)';
  summaryTitleCell.font = { size: 12, bold: true };
  summaryTitleCell.alignment = { horizontal: 'center' };

  const formatCurrency = (value) => value.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' });

  // Adiciona as linhas de dados usando quatro colunas (C, D, E, F)
  const addFinanceRow = (label, value) => {
    const row = worksheet.addRow(['', '', label, '', value, '']);
    // Mescla as colunas C e D para o label
    worksheet.mergeCells(`C${row.number}:D${row.number}`);
    // Mescla as colunas E e F para o valor
    worksheet.mergeCells(`E${row.number}:F${row.number}`);
    return row;
  };

  addFinanceRow('Saldo inicial (para troco)', formatCurrency(valores.saldoInicial));
  addFinanceRow('Total em Espécie', formatCurrency(valores.totalEspecie));
  addFinanceRow('Total em TPA/RUPE', formatCurrency(valores.totalTPA));
  addFinanceRow('Total em Isentos', formatCurrency(valores.totalIsentos));

  const totalGeral = valores.totalEspecie + valores.totalTPA;
  const totalGeralRow = addFinanceRow('Total Geral (Espécie + TPA)', formatCurrency(totalGeral));
  totalGeralRow.getCell(3).font = { bold: true };
  totalGeralRow.getCell(5).font = { bold: true };

  addFinanceRow('Valor Declarado', formatCurrency(valores.valorDeclarado));

  const diferenca = valores.valorDeclarado - totalGeral;
  const diferencaRow = addFinanceRow('Diferença', `${diferenca >= 0 ? '+' : ''}${formatCurrency(diferenca)}`);
  diferencaRow.getCell(3).font = { bold: true };
  diferencaRow.getCell(5).font = { bold: true };

  const endRow = worksheet.lastRow.number;
  
  // Aplica bordas às quatro colunas (C a F)
  applyBordersToRange(worksheet, `${startCol}${startRow}:${endCol}${endRow}`);
  
  // Aplica preenchimento ao título
  applyFill(worksheet, [`${startCol}${startRow}`], 'FFE6E6E6');

  return { totalGeral, diferenca };
};


/**
 * Cria o rodapé com observações e assinaturas.
 * @param {ExcelJS.Worksheet} worksheet - A folha de cálculo.
 */
const criarRodape = (worksheet) => {
    worksheet.addRow([]);
    worksheet.addRow(['Observação:']);
    worksheet.addRow([]);
    worksheet.addRow([]);
    worksheet.addRow(['_______________________________', '', '', '', '_______________________________']);
    worksheet.addRow(['Operador(a)', '', '', '', 'Chefe de Turno']);
    worksheet.addRow([]);
    worksheet.addRow(['X-access portagens - Portagem da barra do Kwanza']);
};


// --- Controlador Principal ---

export async function gerarFechoDeCaixa(req, res) {
  try {
    // 1. Obter e processar dados (aqui usamos mock)
    const dadosMockados = {
      chefeTurno: 'Luciano Alberto',
      cabine: 'Cabine 2 / Pista 2',
      ref: 'REF001',
      dataAbertura: '21-01-2025 00:14:46',
      dataFechamento: '21-01-2025 07:10:53',
      operador: 'Ilidio Pundi',
      veiculos: {
        A: { tarifa: 150.00, especie: 0, tpa: 0, isento: 0 },
        A1: { tarifa: 100.00, especie: 0, tpa: 0, isento: 0 },
        B: { tarifa: 500.00, especie: 21, tpa: 1, isento: 3 },
        C: { tarifa: 2000.00, especie: 38, tpa: 3, isento: 1 },
        C1: { tarifa: 1000.00, especie: 56, tpa: 0, isento: 0 },
      },
      valores: {
        saldoInicial: 0.00,
        totalEspecie: 142500.00,
        totalTPA: 6500.00,
        totalIsentos: 3500.00,
        valorDeclarado: 221500.00,
      },
    };

    // 2. Inicializar Workbook e Worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Fecho de Caixa');

    // 3. Construir o relatório por secções
    criarCabecalho(worksheet, dadosMockados);
    const totaisTabela = criarTabelaVeiculos(worksheet, dadosMockados.veiculos);
    
    // Adicionar linha de totais à tabela
    const totalRow = worksheet.addRow([
        'Total',
        totaisTabela.totalEspecieVeiculos,
        dadosMockados.valores.totalEspecie.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
        totaisTabela.totalTpaVeiculos,
        dadosMockados.valores.totalTPA.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
        totaisTabela.totalIsentoVeiculos,
        dadosMockados.valores.totalIsentos.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
        totaisTabela.totalGeralVeiculos,
        totaisTabela.totalGeralValor.toLocaleString('pt-AO', { style: 'currency', currency: 'AOA' }),
    ]);
    totalRow.font = { bold: true };

    const { totalGeral, diferenca } = criarResumoFinanceiro(worksheet, dadosMockados.valores);
    // criarRodape(worksheet);

    // 4. Aplicar formatação e estilos
    worksheet.columns = [
        { width: 25 }, { width: 15 }, { width: 15 }, { width: 15 },
        { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }
    ];
    
    const tableRange = `A9:I${11 + Object.keys(dadosMockados.veiculos).length}`; // 11 é a linha inicial dos dados
    applyBordersToRange(worksheet, tableRange);
    applyFill(worksheet, ['A9', 'B9', 'D9', 'F9', 'H9'], 'FFE6E6E6'); // Cabeçalho principal
    applyFill(worksheet, ['A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10'], 'FFF0F0F0'); // Sub-cabeçalho

    // 5. Gerar e enviar o ficheiro
    const nomeArquivo = gerarNomeArquivo('fecho-caixa');
    const caminho = path.join(__dirname, '..', '..', 'downloads', nomeArquivo);
    await workbook.xlsx.writeFile(caminho);

    res.json({
      success: true,
      message: 'Relatório de fecho de caixa gerado com sucesso!',
      downloadUrl: `/downloads/${nomeArquivo}`,
      dados: {
        totalVeiculos: totaisTabela.totalGeralVeiculos,
        totalArrecadado: totalGeral,
        diferenca: diferenca,
        operador: dadosMockados.operador,
        periodo: `${dadosMockados.dataAbertura} - ${dadosMockados.dataFechamento}`,
      },
    });
  } catch (error) {
    console.error('Erro ao gerar planilha:', error); // Adicionado para melhor depuração
    res.status(500).json({
      success: false,
      message: 'Erro ao gerar planilha de fecho de caixa',
      error: error.message,
    });
  }
}