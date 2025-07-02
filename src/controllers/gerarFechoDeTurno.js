import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// Supondo que você tenha uma função para gerar nomes de arquivo
const gerarNomeArquivo = (prefixo) => `${prefixo}-${Date.now()}.xlsx`;

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// --- Funções de Utilitário para Excel Melhoradas ---

/**
 * Aplica estilos detalhados a uma célula.
 * @param {ExcelJS.Cell} cell - A célula a ser estilizada.
 * @param {object} styles - Objeto com estilos a aplicar.
 */
const applyCellStyles = (cell, { text, font = {}, alignment = {}, fill = null, border = {}, numFmt = null }) => {
  if (text !== undefined) cell.value = text;
  cell.font = { name: 'Calibri', size: 10, ...font };
  cell.alignment = { vertical: 'middle', ...alignment };
  if (fill) cell.fill = fill;
  if (border) cell.border = border;
  if (numFmt) cell.numFmt = numFmt;
};

/**
 * Aplica bordas a um intervalo de células.
 * @param {ExcelJS.Worksheet} worksheet - A folha de cálculo.
 * @param {string} range - O intervalo de células (ex: 'A1:H10').
 * @param {object} borderStyle - Estilo da borda.
 */
const applyBordersToRange = (worksheet, range, borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }) => {
  const [startCell, endCell] = range.split(':');
  const startRow = parseInt(startCell.match(/\d+/)[0], 10);
  const startCol = worksheet.getColumn(startCell.match(/[A-Z]+/)[0]).number;
  const endRow = parseInt(endCell.match(/\d+/)[0], 10);
  const endCol = worksheet.getColumn(endCell.match(/[A-Z]+/)[0]).number;

  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      worksheet.getCell(r, c).border = { ...worksheet.getCell(r, c).border, ...borderStyle };
    }
  }
};

// --- Funções de Geração de Secções do Relatório ---

const criarCabecalho = (worksheet, dados) => {
  // Logo FROE
  worksheet.mergeCells('A1:C4');
  applyCellStyles(worksheet.getCell('A1'), {
    text: 'FROE\nFundo Rodoviário\ne Obras de Emergência',
    font: { bold: true, size: 10 },
    alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
  });

  // Título Principal
  worksheet.mergeCells('D1:L2');
  applyCellStyles(worksheet.getCell('D1'), {
    text: 'Relatório de fecho de Turno',
    font: { size: 18, bold: true },
    alignment: { horizontal: 'center' },
  });

  // Informações do Turno
  const infoStyle = { font: { bold: true, size: 10 }, alignment: { horizontal: 'left' } };
  applyCellStyles(worksheet.getCell('A5'), { text: 'Chefe de Turno:', ...infoStyle });
  applyCellStyles(worksheet.getCell('B5'), { text: dados.chefeTurno });
  applyCellStyles(worksheet.getCell('E5'), { text: 'Turno:', ...infoStyle });
  applyCellStyles(worksheet.getCell('F5'), { text: dados.turno });
  applyCellStyles(worksheet.getCell('H5'), { text: 'Ref:', ...infoStyle });
  applyCellStyles(worksheet.getCell('I5'), { text: dados.ref });

  applyCellStyles(worksheet.getCell('A6'), { text: 'Data de Abertura:', ...infoStyle });
  worksheet.mergeCells('B6:C6');
  applyCellStyles(worksheet.getCell('B6'), { text: dados.dataAbertura });
  applyCellStyles(worksheet.getCell('A7'), { text: 'Data de Fechamento:', ...infoStyle });
  worksheet.mergeCells('B7:C7');
  applyCellStyles(worksheet.getCell('B7'), { text: dados.dataFechamento });
  
  worksheet.getRow(8).height = 10; // Espaçamento
};

const criarTabelaVeiculos = (worksheet, veiculos) => {
  const startRow = 9;
  // Título da Tabela
  worksheet.mergeCells(`A${startRow}:L${startRow}`);
  applyCellStyles(worksheet.getCell(`A${startRow}`), {
    text: 'REGISTO DE VEÍCULOS & VALOR PAGO',
    font: { size: 12, bold: true },
    alignment: { horizontal: 'center' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } },
  });

  // Cabeçalhos
  const headerRow1 = worksheet.getRow(startRow + 1);
  const headers1 = ['CLASSE', 'ESPÉCIE', null, 'TPA/RUPE', null, 'VIA VERDE', null, 'ISENTO', null, 'TOTAL'];
  headerRow1.values = headers1;
  worksheet.mergeCells('B10:C10');
  worksheet.mergeCells('D10:E10');
  worksheet.mergeCells('F10:G10');
  worksheet.mergeCells('H10:I10');
  worksheet.mergeCells('J10:L10');
  headerRow1.eachCell(cell => applyCellStyles(cell, { font: { bold: true }, alignment: { horizontal: 'center' } }));

  const headerRow2 = worksheet.getRow(startRow + 2);
  headerRow2.values = ['','Nº de\nVeículo','Valor\n(Kz)','Nº de\nVeículo','Valor\n(Kz)','Nº de\nVeículo','Valor\n(Kz)','Nº de\nVeículo','Valor\n(Kz)','Nº de\nVeículo','Valor\n(Kz)'];
  headerRow2.eachCell(cell => applyCellStyles(cell, { font: { bold: true, size: 9 }, alignment: { horizontal: 'center', wrapText: true } }));

  // Dados
  veiculos.forEach((veiculo, index) => {
    const row = worksheet.getRow(startRow + 3 + index);
    const isTotalRow = veiculo[0] === 'Total';
    row.values = veiculo.map((v, i) => (i > 0 && v !== '--' && !isTotalRow) ? Number(v) : v);
    
    row.eachCell((cell, colNumber) => {
        const isCurrency = [3, 5, 7, 9, 11].includes(colNumber);
        const isCount = [2, 4, 6, 8, 10].includes(colNumber);
        applyCellStyles(cell, {
            font: { bold: isTotalRow },
            alignment: { horizontal: 'center' },
            numFmt: isCurrency ? '#,##0' : (isCount ? '0' : undefined)
        });
    });
    applyCellStyles(row.getCell(1), { alignment: { horizontal: 'left' } });
  });

  // Bordas
  const tableRange = `A${startRow}:L${startRow + 2 + veiculos.length}`;
  applyBordersToRange(worksheet, tableRange);
};

const criarValoresDeclarados = (worksheet, valores) => {
    const startRow = worksheet.lastRow.number + 2;
    
    // Título
    worksheet.mergeCells(`H${startRow}:L${startRow}`);
    applyCellStyles(worksheet.getCell(`H${startRow}`), {
        text: 'VALORES DECLARADOS',
        font: { bold: true },
        alignment: { horizontal: 'center' },
    });
    applyCellStyles(worksheet.getCell(`L${startRow}`), { text: 'AKZ', font: { bold: true }, alignment: { horizontal: 'center' } });

    // Dados
    valores.forEach(([item, valor], index) => {
        const currentRow = startRow + 1 + index;
        const isTotal = item.toLowerCase().includes('total');
        worksheet.mergeCells(`H${currentRow}:K${currentRow}`);
        applyCellStyles(worksheet.getCell(`H${currentRow}`), {
            text: item,
            font: { bold: isTotal },
            alignment: { horizontal: 'left' }
        });
        applyCellStyles(worksheet.getCell(`L${currentRow}`), {
            text: Number(valor.replace(/[^0-9,]/g, '').replace(',', '.')),
            font: { bold: isTotal },
            alignment: { horizontal: 'right' },
            numFmt: '"kz" #,##0.00'
        });
    });

    // Bordas
    applyBordersToRange(worksheet, `H${startRow}:L${startRow + valores.length}`);
};

const criarTabelaExcedentes = (worksheet, excedentes) => {
    const startRow = worksheet.lastRow.number - excedentes.length; // Alinhar com a tabela de valores
    
    // Título
    worksheet.mergeCells(`A${startRow}:E${startRow}`);
    applyCellStyles(worksheet.getCell(`A${startRow}`), {
        text: 'DECLARAÇÃO DE EXCEDENTES',
        font: { bold: true },
        alignment: { horizontal: 'center' },
    });

    // Cabeçalho
    const headerRow = worksheet.getRow(startRow + 1);
    worksheet.mergeCells(`A${startRow + 1}:D${startRow + 1}`);
    applyCellStyles(headerRow.getCell('A'), { text: 'Nome', font: { bold: true } });
    applyCellStyles(headerRow.getCell('E'), { text: 'Diferença', font: { bold: true }, alignment: { horizontal: 'right' } });

    // Dados
    excedentes.forEach(([nome, diferenca], index) => {
        const currentRow = startRow + 2 + index;
        worksheet.mergeCells(`A${currentRow}:D${currentRow}`);
        applyCellStyles(worksheet.getCell(`A${currentRow}`), { text: nome, alignment: { horizontal: 'left' } });
        applyCellStyles(worksheet.getCell(`E${currentRow}`), {
            text: Number(diferenca.replace(/[^0-9,-]/g, '').replace(',', '.')),
            alignment: { horizontal: 'right' },
            numFmt: '#,##0.00'
        });
    });

    // Bordas
    applyBordersToRange(worksheet, `A${startRow}:E${startRow + 1 + excedentes.length}`);
};

const criarDescricaoEAssinaturas = (worksheet, descricao) => {
    let startRow = worksheet.lastRow.number + 3;

    applyCellStyles(worksheet.getCell(`A${startRow}`), { text: 'Descrição da Actividade', font: { bold: true } });
    
    startRow++;
    worksheet.mergeCells(`A${startRow}:L${startRow + 2}`);
    applyCellStyles(worksheet.getCell(`A${startRow}`), {
        text: descricao,
        alignment: { wrapText: true, vertical: 'top' }
    });

    startRow += 5; // Espaço para assinaturas
    applyCellStyles(worksheet.getCell(`A${startRow}`), { text: '_______________________', alignment: { horizontal: 'center' } });
    applyCellStyles(worksheet.getCell(`E${startRow}`), { text: '_______________________', alignment: { horizontal: 'center' } });
    applyCellStyles(worksheet.getCell(`J${startRow}`), { text: '_______________________', alignment: { horizontal: 'center' } });
    
    startRow++;
    applyCellStyles(worksheet.getCell(`A${startRow}`), { text: 'Chefe de turno', alignment: { horizontal: 'center' } });
    applyCellStyles(worksheet.getCell(`E${startRow}`), { text: 'Superv. ADM. FIN', alignment: { horizontal: 'center' } });
    applyCellStyles(worksheet.getCell(`J${startRow}`), { text: 'Coordenador Geral', alignment: { horizontal: 'center' } });

    // Rodapé
    startRow += 3;
    const geradoEm = new Date().toLocaleString('pt-AO', { dateStyle: 'short', timeStyle: 'short' });
    applyCellStyles(worksheet.getCell(`A${startRow}`), { text: `Ref: 0013\ngerado em: ${geradoEm}`, font: { size: 8 }, alignment: { wrapText: true } });
    applyCellStyles(worksheet.getCell(`F${startRow}`), { text: 'X-access - Portagem da barra do Kwanza', font: { size: 8 }, alignment: { horizontal: 'center' } });
    applyCellStyles(worksheet.getCell(`L${startRow}`), { text: 'Página 1', font: { size: 8 }, alignment: { horizontal: 'right' } });
};

// --- Controlador Principal ---

export async function gerarFechoDeTurno(req, res) {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Fecho de Turno');

    // --- DADOS MOCKADOS ---
    const dados = {
      chefeTurno: 'Bartolomeu zeca',
      turno: 5,
      ref: '0013',
      dataAbertura: '30-01-2025 07:21:27',
      dataFechamento: '31-01-2025 07:32:46',
    };
    const veiculos = [
      ['A (150)', 3, 450, 0, '--', 0, '--', 0, '--', 3, 450],
      ['A1 (100)', 1, 100, 0, '--', 0, '--', 0, '--', 1, 100],
      ['B (500)', 780, 390000, 39, 19500, 0, '--', 12, 6000, 831, 415500],
      ['C (2000)', 347, 694000, 16, 32000, 37, 74000, 8, 16000, 371, 742000],
      ['C1 (1000)', 283, 283000, 2, 2000, 0, '--', 9, 9000, 294, 294000],
      ['Total', 1414, 1367550, 57, 53500, 37, 74000, 29, 31000, 1500, 1452050],
    ];
    const valoresDeclarados = [
      ['Total em Espécie', '1367550,00'],
      ['Total em TPA/RUPE', '53500,00'],
      ['Total em Via Verde', '74000,00'],
      ['Total em Isentos', '31000,00'],
      ['Total Geral(Espécie + TPA)', '1421050,00'],
      ['Total Declarado', '2164950,00'],
      ['Total diferença', '+900,00'],
    ];
    const excedentes = [
      ['Amaro Tomas', '0,00'], ['Augusto Agostinho', '-2000,00'], ['Anselmo Luis', '+1600,00'],
      ['Braulio Silva', '0,00'], ['Moises Gouveia', '+900,00'], ['Manuel Kali', '-100,00'],
      ['Osvaldo Delgado', '+500,00'], ['Mbozo Ricardo', '0,00'], ['Domingos Boa', '0,00'],
    ];
    const descricaoAtividade = 'Dizer que não tivemos anomalias durante o turno, quer do ponto de vista dos meios técnicos, quer a nível do pessoal, resultando assim; numa jornada laboral produtiva e tranquila.';

    // --- CONSTRUIR O RELATÓRIO ---
    criarCabecalho(worksheet, dados);
    criarTabelaVeiculos(worksheet, veiculos);
    criarValoresDeclarados(worksheet, valoresDeclarados);
    criarTabelaExcedentes(worksheet, excedentes);
    criarDescricaoEAssinaturas(worksheet, descricaoAtividade);

    // Configurar largura das colunas
    worksheet.columns = [
        { width: 12 }, { width: 8 }, { width: 12 }, { width: 8 }, { width: 12 },
        { width: 8 }, { width: 12 }, { width: 8 }, { width: 12 }, { width: 8 },
        { width: 12 }, { width: 12 }
    ];
    worksheet.getColumn('A').width = 20;
    worksheet.getColumn('H').width = 10;
    worksheet.getColumn('I').width = 10;
    worksheet.getColumn('J').width = 10;
    worksheet.getColumn('K').width = 10;
    worksheet.getColumn('L').width = 15;


    // --- GERAR E SALVAR O ARQUIVO ---
    const nomeArquivo = gerarNomeArquivo('fecho-turno-melhorado');
    const caminho = path.join(__dirname, '..', '..', 'downloads', nomeArquivo); // Ajuste o caminho conforme sua estrutura
    await workbook.xlsx.writeFile(caminho);

    // Enviar resposta
    res.json({
      success: true,
      message: 'Relatório de fecho de turno melhorado gerado com sucesso!',
      downloadUrl: `/downloads/${nomeArquivo}`,
    });
  } catch (error) {
    console.error('Erro ao gerar relatório de fecho de turno:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao gerar relatório de fecho de turno',
      error: error.message,
    });
  }
}
