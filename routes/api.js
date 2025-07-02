import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const router = express.Router();

// Rota: Listar templates disponíveis
router.get('/templates', (req, res) => {
  const templates = [
    {
      id: 'basic',
      nome: 'Planilha Básica',
      descricao: 'Planilha simples com dados de funcionários',
      endpoint: '/excel/basic',
      metodo: 'GET'
    },
    {
      id: 'vendas',
      nome: 'Relatório de Vendas',
      descricao: 'Relatório completo com fórmulas e análises',
      endpoint: '/excel/vendas',
      metodo: 'GET'
    },
    {
      id: 'produtos',
      nome: 'Lista de Produtos',
      descricao: 'Catálogo de produtos com cálculos de estoque',
      endpoint: '/excel/produtos',
      metodo: 'GET'
    },
    {
      id: 'custom',
      nome: 'Planilha Personalizada',
      descricao: 'Crie sua própria planilha enviando dados via POST',
      endpoint: '/excel/custom',
      metodo: 'POST',
      parametros: {
        titulo: 'string - Título da planilha',
        colunas: 'array - Definição das colunas [{header, key, width}]',
        dados: 'array - Dados para preencher a planilha'
      }
    }
  ];

  res.json({
    success: true,
    templates,
    total: templates.length
  });
});

// Rota: Listar arquivos disponíveis para download
router.get('/downloads', (req, res) => {
  try {
    const downloadsDir = path.join(__dirname, '..', 'downloads');
    const arquivos = fs.readdirSync(downloadsDir)
      .filter(arquivo => arquivo.endsWith('.xlsx'))
      .map(arquivo => {
        const stats = fs.statSync(path.join(downloadsDir, arquivo));
        return {
          nome: arquivo,
          tamanho: stats.size,
          criado: stats.birthtime,
          downloadUrl: `/downloads/${arquivo}`
        };
      })
      .sort((a, b) => b.criado - a.criado); // Mais recentes primeiro

    res.json({
      success: true,
      arquivos,
      total: arquivos.length
    });

  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Erro ao listar arquivos',
      error: error.message
    });
  }
});

// Rota: Informações do sistema
router.get('/info', (req, res) => {
  res.json({
    success: true,
    sistema: {
      nome: 'ExcelJS Web Server',
      versao: '1.0.0',
      node: process.version,
      uptime: process.uptime(),
      memoria: process.memoryUsage()
    },
    rotas: {
      excel: [
        'GET /excel/basic',
        'GET /excel/vendas',
        'GET /excel/produtos',
        'POST /excel/custom'
      ],
      api: [
        'GET /api/templates',
        'GET /api/downloads',
        'GET /api/info'
      ]
    }
  });
});

export default router;
