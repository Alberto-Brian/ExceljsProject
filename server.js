import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// Simular __dirname para ES Modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Importar rotas (com extensão .js obrigatória)
import excelRoutes from './routes/excel.js';
import apiRoutes from './routes/api.js';

const app = express();
const PORT = process.env.PORT || 3000;

// Middlewares
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Servir arquivos estáticos
app.use(express.static('public'));
app.use('/downloads', express.static('downloads'));

// Criar pasta de downloads se não existir
const downloadsDir = path.join(__dirname, 'downloads');
if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir);
  console.log('✅ Pasta ./downloads/ criada');
}

// Rotas
app.use('/excel', excelRoutes);
app.use('/api', apiRoutes);

// Página principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Middleware de erros
app.use((err, req, res, next) => {
  console.error('Erro:', err.message);
  res.status(500).json({
    success: false,
    message: 'Erro interno do servidor',
    error: err.message
  });
});

// Iniciar servidor
app.listen(PORT, () => {
  console.log('🚀 Servidor ExcelJS iniciado!');
  console.log(`📡 Rodando em: http://localhost:${PORT}`);
  console.log('📊 Rotas disponíveis:');
  console.log('  GET  /                    - Interface web');
  console.log('  GET  /excel/basic         - Gerar planilha básica');
  console.log('  GET  /excel/vendas        - Gerar relatório de vendas');
  console.log('  GET  /excel/produtos      - Gerar lista de produtos');
  console.log('  POST /excel/custom        - Gerar planilha personalizada');
  console.log('  GET  /api/templates       - Listar templates disponíveis');
  console.log('  GET  /downloads           - Arquivos para download');
  console.log('\n💡 Acesse http://localhost:3000 para começar!');
});
