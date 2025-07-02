# 🚀 ExcelJS Web Server

Servidor web completo para gerar arquivos Excel usando ExcelJS com API REST e interface web.

## 📋 Funcionalidades

### 🌐 Interface Web
- Interface moderna e responsiva
- Geração de planilhas com um clique
- Formulário para planilhas personalizadas
- Lista de downloads em tempo real
- Documentação da API integrada

### 📊 Templates de Excel
- **Planilha Básica**: Dados de funcionários com formatação
- **Relatório de Vendas**: Fórmulas Excel e análises avançadas
- **Lista de Produtos**: Cálculos de estoque e valores
- **Planilha Personalizada**: Envie seus próprios dados via API

### 🔗 API REST Completa
- Endpoints para gerar diferentes tipos de planilhas
- Suporte a dados personalizados via POST
- Listagem de templates e arquivos
- Informações do sistema

## 🚀 Como usar

### Instalação e Execução
```bash
# Instalar dependências
npm install

# Iniciar servidor
npm start

# Ou modo desenvolvimento (com nodemon)
npm run dev
```

### Acessar a aplicação
```
http://localhost:3000
```

## 📡 API Endpoints

### Gerar Planilhas Excel
```http
GET  /excel/basic         # Planilha básica
GET  /excel/vendas        # Relatório de vendas  
GET  /excel/produtos      # Lista de produtos
POST /excel/custom        # Planilha personalizada
```

### Informações e Utilitários
```http
GET /api/templates        # Listar templates
GET /api/downloads        # Listar arquivos
GET /api/info            # Info do sistema
GET /downloads           # Pasta de arquivos
```

## 📝 Exemplo de Uso da API

### Gerar Planilha Básica
```javascript
const response = await fetch('http://localhost:3000/excel/basic');
const data = await response.json();

console.log(data);
// {
//   "success": true,
//   "message": "Planilha básica gerada com sucesso!",
//   "arquivo": "planilha-basica-2024-01-15T10-30-00.xlsx",
//   "downloadUrl": "/downloads/planilha-basica-2024-01-15T10-30-00.xlsx",
//   "registros": 4
// }
```

### Gerar Planilha Personalizada
```javascript
const dadosPersonalizados = {
  titulo: "Minha Planilha",
  colunas: [
    { header: "ID", key: "id", width: 10 },
    { header: "Nome", key: "nome", width: 20 },
    { header: "Email", key: "email", width: 30 }
  ],
  dados: [
    { id: 1, nome: "João Silva", email: "joao@email.com" },
    { id: 2, nome: "Maria Santos", email: "maria@email.com" }
  ]
};

const response = await fetch('http://localhost:3000/excel/custom', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(dadosPersonalizados)
});

const resultado = await response.json();
```

## 🛠️ Recursos do ExcelJS Demonstrados

### ✅ Criação e Estrutura
- Workbooks e Worksheets múltiplas
- Configuração de colunas com largura
- Adição de dados estruturados
- Mesclagem de células

### 🎨 Formatação Avançada
- Fontes personalizadas (negrito, cores)
- Cores de fundo e padrões
- Alinhamento de texto
- Formatação de números (moeda, percentual)
- Bordas e estilos

### 📊 Fórmulas e Cálculos
- Fórmulas Excel nativas (SUM, AVERAGE, IF)
- Referências entre células
- Cálculos automáticos
- Funções condicionais

### 🔧 Recursos Técnicos
- Geração de nomes únicos com timestamp
- Tratamento de erros robusto
- Validação de dados de entrada
- Logs detalhados

## 📁 Estrutura do Projeto

```
├── server.js              # Servidor principal Express
├── routes/
│   ├── excel.js          # Rotas para gerar Excel
│   └── api.js            # Rotas de informações
├── public/
│   └── index.html        # Interface web
├── downloads/            # Arquivos Excel gerados
└── README.md
```

## 🔧 Tecnologias Utilizadas

- **Node.js** - Runtime JavaScript
- **Express.js** - Framework web
- **ExcelJS** - Biblioteca para Excel
- **CORS** - Suporte a requisições cross-origin
- **HTML/CSS/JS** - Interface web moderna

## 💡 Casos de Uso

### Para Desenvolvedores
- Integrar geração de Excel em aplicações web
- Criar relatórios automatizados
- Exportar dados de APIs para Excel
- Prototipagem rápida de funcionalidades

### Para Empresas
- Relatórios gerenciais automatizados
- Exportação de dados de sistemas
- Geração de planilhas padronizadas
- Integração com sistemas existentes

## 🚀 Próximos Passos

Após dominar este servidor, você pode:
- Integrar com bancos de dados
- Adicionar autenticação e autorização
- Implementar templates mais complexos
- Adicionar gráficos e imagens
- Criar sistema de agendamento de relatórios
- Implementar cache de arquivos
- Adicionar validação de dados mais robusta

---

**🌟 Acesse http://localhost:3000 e comece a gerar planilhas Excel!**